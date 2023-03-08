<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/MD5.asp"-->
<%
Dim rs,logid,rstmp,password,logfile,uid,blogpw,action,encommment,commenttopic,log_month,logtopic
Dim show,blog,authorid
Dim SQL
logid=CLng(Request("id"))
password=Trim(Request("password"))
action=Trim(Request("action"))
G_P_This=CLng(Request("page"))
G_P_PerMax=10
If logid=0 Then
	oblog.adderrstr("日志参数错误")
	oblog.showerr
End If
SQL = " logfile,userid,topic,ishide,ispassword,blog_password,isneedlogin,viewscores,IsSpecial,viewgroupid ,isencomment,addtime"
Set rs=oblog.execute("select "&SQL&"  from oblog_log where logid="&logid)
If rs.eof Then
	Set rs=Nothing
	oblog.adderrstr("无此日志")
	oblog.showerr
End If
logfile=rs(0)
uid=rs(1)
logtopic=rs(2)
'If rs(8) = 1 Or IsNull(rs(8)) Then
	'Set rs=Nothing
	'Response.Redirect(logfile)
'End If

'加密博客
If rs(5)=1 Then
	blogpw=Request.Cookies(cookies_name)("blog_pwd_"&uid)
	Set rstmp=oblog.execute("select * from oblog_user where userid="&uid)
	If (rstmp("blog_password")<>"" or IsNull(rstmp("blog_password"))=False) And blogpw<>rstmp("blog_password") Then
		Set rs=Nothing
		Set rstmp=Nothing
		Response.Redirect("chkblogpassword.asp?userid="&uid&"&fromurl="&Replace(oblog.GetUrl,"&","$"))
	End If
'隐藏日志
ElseIf rs(3)=1 Then
	If not oblog.checkuserlogined() Then
		oblog.adderrstr("需要登录才可以查看隐藏日志!")
		oblog.showerr
	else
		Set rstmp=oblog.execute("select id from oblog_friEnd where userid="&uid&" And friEndid="&oblog.l_uid)
		If rstmp.eof And oblog.l_uid<>uid Then
			Set rs=Nothing
			Set rstmp=Nothing
			oblog.adderrstr("您无权限查看此日志，请联系blog主人!")
			oblog.showerr
		else
			Set rstmp=Nothing
		End If
	End If
'加密日志
ElseIf rs(4)<>"" Then
	If password="" And Request.Cookies(cookies_name)("logpw_"&logid)="" And action="" Then
		Set rs=Nothing
		Response.Redirect(logfile)
	End If
	If password<>"" Then password=MD5(password)
	If password<>rs(4) And  Request.Cookies(cookies_name)("logpw_"&logid)<>rs(4) Then
		Set rs=Nothing
		oblog.adderrstr("日志访问密码错误,请重新输入!")
		oblog.showerr
	else
		If password<>"" Then Response.Cookies(cookies_name)("logpw_"&logid)=password
		Set rs=Nothing
	End If
'积分浏览，登录可见，用户组可见
ElseIf OB_IIF(rs(6),0) = 1 Or OB_IIF(rs(7),0) > 0 Or OB_IIF(rs(9),0) > 0 Then
	If Not Oblog.CheckUserLogined Then
		oblog.adderrstr("您无权查看此日志，请登录后查看!")
		oblog.showerr
	End If
	If OB_IIF(rs(7),0) > 0 Then
		If Oblog.l_uid <> uid Then
			If oblog.CheckScore(rs(7)) Then
				oblog.GiveScore "",-1*Abs(rs(7)),Oblog.l_uid
			Else
				oblog.adderrstr("您无权查看此日志，您的积分不足!")
				oblog.showerr
			End if
			oblog.GiveScore "",rs(7),uid
		End if
	ElseIf OB_IIF(rs(9),0) > 0 Then
		If Oblog.l_uid <> uid Then
			If Oblog.l_uGroupId <> rs(9) Then
				oblog.adderrstr("您无权查看此日志，您所在的用户组不匹配!")
				oblog.showerr
			End If
		End if
	End If
'隐藏日志（一般用户个人分类下的日志隐藏）
Else
	If OB_IIF(rs(8),0) > 0 Then
		If not oblog.checkuserlogined() Then
			oblog.adderrstr("非此blog主人无权查看此日志!")
			oblog.showerr
		else
			If oblog.l_uid<>uid Then
				Set rs=Nothing
				Set rstmp=Nothing
				oblog.adderrstr("您无权限查看此日志，请联系blog主人!")
				oblog.showerr
			else
				Set rstmp=Nothing
			End If
		End If
	End if
End If

call main()

Sub main()
	select Case action
		Case "comment"
			call showcomment
		Case else
			call showlog()
	End select
End Sub

Sub showlog()
	Set blog=new class_blog
	blog.userid=uid
	blog.showpwlog=True
	blog.update_log logid,0
	show=blog.filt_pwblog(blog.m_log,logtopic)
	show=Replace(show,"savecomment.asp?logid=","savecomment.asp?t=1&logid=")
	Response.Write(show)
	Set blog=Nothing
End Sub

Sub showcomment()
	Set blog=new class_blog
	encommment=rs("isencomment")
	commenttopic="Re:"&oblog.filt_html(logtopic)
	If Int(Month(rs("addtime")))<10 Then
		log_month=Year(rs("addtime"))&"0"&Month(rs("addtime"))
	else
		log_month=Year(rs("addtime"))&Month(rs("addtime"))
	End If
	blog.userid=uid
	blog.showpwlog = True
	blog.showcmt logid
	show=blog.filt_pwblog(blog.m_commentsmore,rs("topic")&"--所有评论")
	Response.Write(show)
	Set blog=Nothing
End Sub
%>