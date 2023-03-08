<!--#include file="inc/inc_syssite.asp"-->
<%
dim rs,uname,domain,sql,hideurl,subjectid,reurl,uid
uname = oblog.filt_badstr(Request("name"))
subjectid = CLng (Request("subjectid"))
domain = Trim (Request("domain"))
if uname="" and domain="" then
	oblog.adderrstr("参数错误")
	oblog.showerr
end if
if domain<>"" then
	dim domain1,domain2
	domain = LCase (domain)
	domain = Replace (domain,"http://","")
	domain = Replace (domain,"/","")
	domain1= oblog.filt_badstr(Replace (Left (domain,InStr (domain,".")),".",""))
	if Trim (domain1)="" Then Response.Redirect("index.asp"):Response.End
	domain2= oblog.filt_badstr(Right(domain,Len(domain)-InStr(domain,".")))
	If Oblog.CheckDomainRoot(domain2,0) Then
		sql = "select user_dir,userid,hideurl,user_folder from oblog_user where user_domain='"&domain1&"' and user_domainroot='"&domain2&"'"
	ElseIf Oblog.CheckDomainRoot(domain2,1) Then
		Call TeamDomain()
	End if
Elseif uname<>"" then
	sql="select user_dir,userid,hideurl,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_user where username='"&uname&"'"
end If

If sql = "" Then
	oblog.adderrstr("域名根不合法")
	oblog.showerr
End if
set rs=oblog.execute(sql)
if not rs.eof then
	hideurl=rs(2)
	if true_domain=1 then
		if rs("custom_domain")="" or isnull(rs("custom_domain")) then
			reurl="http://"&rs("user_domain")&"."&rs("user_domainroot")
		else
			reurl="http://"&rs("custom_domain")
		end if
	else
		reurl=blogdir&rs(0)&"/"&rs(3)
	end if
	if subjectid>0 then
		reurl=reurl&"/cmd."&f_ext&"?uid="&rs(1)&"&do=blogs&id="&subjectid
	else
		if oblog.cacheConfig(46)=1 then
			reurl=reurl&"/index."&f_ext
		else
			reurl=reurl
		end if
	end if
	set rs=nothing
	if domain<>"" and hideurl=1 then
		Response.Redirect(reurl)
	else
		Response.Write("<script language=JavaScript>top.location='"&reurl&"';</script>")
	end if
else
	set rs=nothing
	oblog.adderrstr("错误：无此blog用户!")
	oblog.showerr
end If
Sub TeamDomain()
	sql = "select teamid,hideurl FROM oblog_team WHERE t_domain = '"&domain1&"' AND t_domainroot='"&domain2&"'"
	set rs=oblog.execute(sql)
	If Not rs.EOF Then
		If rs(1) = 1 Then
			Response.Redirect "group.asp?gid="&rs(0)
		Else
			Response.Write("<script language=JavaScript>top.location='group.asp?gid="&rs(0)&"';</script>")
		End if
	Else
		oblog.adderrstr("错误："&oblog.CacheConfig(69)&"不存在!")
		oblog.showerr
	End If
	Response.End
End Sub
%>