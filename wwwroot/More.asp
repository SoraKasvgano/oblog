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
	oblog.adderrstr("��־��������")
	oblog.showerr
End If
SQL = " logfile,userid,topic,ishide,ispassword,blog_password,isneedlogin,viewscores,IsSpecial,viewgroupid ,isencomment,addtime"
Set rs=oblog.execute("select "&SQL&"  from oblog_log where logid="&logid)
If rs.eof Then
	Set rs=Nothing
	oblog.adderrstr("�޴���־")
	oblog.showerr
End If
logfile=rs(0)
uid=rs(1)
logtopic=rs(2)
'If rs(8) = 1 Or IsNull(rs(8)) Then
	'Set rs=Nothing
	'Response.Redirect(logfile)
'End If

'���ܲ���
If rs(5)=1 Then
	blogpw=Request.Cookies(cookies_name)("blog_pwd_"&uid)
	Set rstmp=oblog.execute("select * from oblog_user where userid="&uid)
	If (rstmp("blog_password")<>"" or IsNull(rstmp("blog_password"))=False) And blogpw<>rstmp("blog_password") Then
		Set rs=Nothing
		Set rstmp=Nothing
		Response.Redirect("chkblogpassword.asp?userid="&uid&"&fromurl="&Replace(oblog.GetUrl,"&","$"))
	End If
'������־
ElseIf rs(3)=1 Then
	If not oblog.checkuserlogined() Then
		oblog.adderrstr("��Ҫ��¼�ſ��Բ鿴������־!")
		oblog.showerr
	else
		Set rstmp=oblog.execute("select id from oblog_friEnd where userid="&uid&" And friEndid="&oblog.l_uid)
		If rstmp.eof And oblog.l_uid<>uid Then
			Set rs=Nothing
			Set rstmp=Nothing
			oblog.adderrstr("����Ȩ�޲鿴����־������ϵblog����!")
			oblog.showerr
		else
			Set rstmp=Nothing
		End If
	End If
'������־
ElseIf rs(4)<>"" Then
	If password="" And Request.Cookies(cookies_name)("logpw_"&logid)="" And action="" Then
		Set rs=Nothing
		Response.Redirect(logfile)
	End If
	If password<>"" Then password=MD5(password)
	If password<>rs(4) And  Request.Cookies(cookies_name)("logpw_"&logid)<>rs(4) Then
		Set rs=Nothing
		oblog.adderrstr("��־�����������,����������!")
		oblog.showerr
	else
		If password<>"" Then Response.Cookies(cookies_name)("logpw_"&logid)=password
		Set rs=Nothing
	End If
'�����������¼�ɼ����û���ɼ�
ElseIf OB_IIF(rs(6),0) = 1 Or OB_IIF(rs(7),0) > 0 Or OB_IIF(rs(9),0) > 0 Then
	If Not Oblog.CheckUserLogined Then
		oblog.adderrstr("����Ȩ�鿴����־�����¼��鿴!")
		oblog.showerr
	End If
	If OB_IIF(rs(7),0) > 0 Then
		If Oblog.l_uid <> uid Then
			If oblog.CheckScore(rs(7)) Then
				oblog.GiveScore "",-1*Abs(rs(7)),Oblog.l_uid
			Else
				oblog.adderrstr("����Ȩ�鿴����־�����Ļ��ֲ���!")
				oblog.showerr
			End if
			oblog.GiveScore "",rs(7),uid
		End if
	ElseIf OB_IIF(rs(9),0) > 0 Then
		If Oblog.l_uid <> uid Then
			If Oblog.l_uGroupId <> rs(9) Then
				oblog.adderrstr("����Ȩ�鿴����־�������ڵ��û��鲻ƥ��!")
				oblog.showerr
			End If
		End if
	End If
'������־��һ���û����˷����µ���־���أ�
Else
	If OB_IIF(rs(8),0) > 0 Then
		If not oblog.checkuserlogined() Then
			oblog.adderrstr("�Ǵ�blog������Ȩ�鿴����־!")
			oblog.showerr
		else
			If oblog.l_uid<>uid Then
				Set rs=Nothing
				Set rstmp=Nothing
				oblog.adderrstr("����Ȩ�޲鿴����־������ϵblog����!")
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
	show=blog.filt_pwblog(blog.m_commentsmore,rs("topic")&"--��������")
	Response.Write(show)
	Set blog=Nothing
End Sub
%>