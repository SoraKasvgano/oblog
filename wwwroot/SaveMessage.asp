<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/inc_antispam.asp"-->
<!--#include file="inc/md5.asp"-->
<%
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
If Application(cache_name_user&"_systemenmod")<>"" Then
	Dim enStr
	enStr=Application(cache_name_user&"_systemenmod")
	enStr=Split(enStr,",")
	If enStr(0)="1" Then	Response.write("系统临时禁止留言!"):Response.End()
End If
if oblog.ChkPost()=false then Response.write("不允许从外部提交"):Response.End()
oblog.chk_commenttime
dim userid,rs,username,password,blog,isguest,message,mainuserid,messagetopic,homepage,ishide,sCheck,checkuserlogined
userid=CLng(Request("userid"))
username=Trim(Request.form("username"))
message=Trim(Request.form("oblog_edittext"))
messagetopic=Trim(Request.Form("commenttopic"))
ishide=Trim(Request.Form("ishide"))
homepage=oblog.InterceptStr(Request.Form("homepage"),250)
if username="" or oblog.strLength(username)>20 then oblog.adderrstr("名字不能为空且不能大于20个字符！")
if oblog.chk_badword(username)>0 then oblog.adderrstr("名字中含有系统不允许的字符！")
if message="" or oblog.strLength(message)>Int(oblog.CacheConfig(35)) then oblog.adderrstr("留言内容不能为空且不能大于"&oblog.CacheConfig(35)&"英文字符)！")
if oblog.chk_badword(message)>0 then oblog.adderrstr("留言内容中含有系统不允许的字符！")
if messagetopic="" or oblog.strLength(messagetopic)>200 then oblog.adderrstr("留言标题不能为空且不能大于200英文字符)！")
if oblog.chk_badword(messagetopic)>0 then oblog.adderrstr("留言标题中含有系统不允许的字符！")
if oblog.chk_badword(homepage)>0 then oblog.adderrstr("主页连接中含有系统不允许的字符！")
'If C_SEditor_HTML=0 Then message=RemoveUBB(RemoveHtml(message))
If ChkCommentTag(message)=false Then
	sCheck=antiSpam("2")
	If sCheck<>"" Then oblog.adderrstr(sCheck)
Else
	 oblog.adderrstr("回复内容中含有系统不允许的字符")
End If
if oblog.errstr<>"" then oblog.ShowMsg Replace(oblog.errstr,"_","\n"),"back"
message=EncodeJP(oblog.filtpath(oblog.filt_badword(message)))
isguest=1
password=Trim(Request.form("password"))
if ishide<>"" then ishide=cint(ishide) else ishide=0
if oblog.CacheConfig(27)=0 or password<>"" then
		password=md5(password)
		oblog.ob_chklogin username,password,0	
end if
checkuserlogined = oblog.checkuserlogined
If checkuserlogined Then isguest=0


if oblog.CacheConfig(27)=0 Then
	If Not checkuserlogined then
		oblog.ShowMsg "需要登录后才能发表评论","back"
		Response.End
	Else
		If oblog.l_ulevel="6" Then
			oblog.ShowMsg "您没有通过管理员审核，不能发表评论","back"
			Response.End()
		End If
	End If
End If

if oblog.CacheConfig(30)=1 Then
	If  Request("CodeStr")="" then
		oblog.ShowMsg "验证码错误，请重新输入","back"
		Response.End()
	Else
		if not oblog.codepass then oblog.ShowMsg "验证码错误，请重新输入","":Response.End()
	End If
end if
if checkuserlogined And password = "" Then username = oblog.l_uname:isguest=0
'Process...
set blog=new class_blog
if not IsObject(conn) then link_database
set rs=Server.CreateObject("adodb.recordset")
rs.open "select top 1 * from oblog_message",conn,2,2
rs.addnew
rs("userid")=userid
rs("message_user")=EncodeJP(oblog.filt_badword(username))
rs("messagetopic")=EncodeJP(oblog.InterceptStr(oblog.filt_badword(messagetopic),250))
rs("message")=message
rs("homepage")=homepage
rs("addtime")=oblog.ServerDate(now())
rs("addip")=oblog.userip
rs("isguest")=isguest
rs("ishide")=ishide
rs("istate")=oblog.CacheConfig(50)
rs.update
rs.close
set rs=Nothing
Dim scores
If oblog.CacheConfig(50) = 0 Then
	scores=0
Else
	scores=oblog.CacheScores(5)
End if
oblog.execute("update oblog_user set lastmessage='" & oblog.ServerDate(Now()) &"' where userid="&userid)
Response.Cookies(cookies_name)("LastComment") = oblog.ServerDate(now())
If oblog.CacheConfig(50)=0 Then
	Set blog=Nothing
	oblog.ShowMsg "留言成功，审核后可见", ""
Else
	oblog.execute("update oblog_user set message_count=message_count+1 ,scores=scores+" & scores&" where userid="&userid)
	oblog.execute("update oblog_setup set message_count=message_count+1")
	blog.userid=userid
	blog.update_message 3
	blog.update_newmessage userid
	Dim GoUrl
	GoUrl=blog.gourl
	Set blog=Nothing
	Response.Redirect GoUrl
End if
%>