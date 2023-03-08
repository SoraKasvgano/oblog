<!--#include file="inc/inc_syssite.asp"-->
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
	If enStr(0)="1" Then Response.write("系统临时禁止评论!"):Response.End()
End If
if oblog.ChkPost()=false then Response.write("不允许从外部提交!"):Response.End()
oblog.chk_commenttime
dim fileid,rs,username,password,blog,isguest,comment,mainuserid,commenttopic,sCheck,teamID,checkuserlogined
fileid=CLng(Request("fileid"))
teamID=Request("teamid")
If teamID <>"" Then teamID = CLng(teamID) Else teamID = 0
username=oblog.filt_badstr(Trim(Request.form("username")))
comment=Trim(Request.form("oblog_edittext"))
commenttopic=Trim(Request.Form("commenttopic"))
if username="" or oblog.strLength(username)>20 then oblog.adderrstr("名字不能为空且不能大于20个字符！")
if oblog.chk_badword(username)>0 then oblog.adderrstr("名字中含有系统不允许的字符！")
if comment="" or oblog.strLength(comment)>Int(oblog.CacheConfig(35)) then oblog.adderrstr("回复内容不能为空且不能大于"&oblog.CacheConfig(35)&"英文字符！")
if oblog.chk_badword(comment)>0 then oblog.adderrstr("回复内容中含有系统不允许的字符！")
If commenttopic="" Then
	If teamID = 0 Then
		oblog.adderrstr("回复标题不能为空！")
	End If
End if
if oblog.strLength(commenttopic)>200 then oblog.adderrstr("回复标题不能不能大于200英文字符！")
if oblog.chk_badword(commenttopic)>0 then oblog.adderrstr("回复标题中含有系统不允许的字符！")
if oblog.chk_badword(Request.Form("homepage"))>0 then oblog.adderrstr("主页地址中含有系统不允许的字符！")
If ChkCommentTag(comment)=false Then
	sCheck=antiSpam("1")
	If sCheck<>"" Then oblog.adderrstr(sCheck)
Else
	 oblog.adderrstr("回复内容中含有系统不允许的字符")
End If
if oblog.errstr<>"" then oblog.ShowMsg Replace(oblog.errstr,"_","\n"),"back":Response.End()
comment= EncodeJP(oblog.filtpath(oblog.filt_badword(comment)))
if oblog.errstr<>"" then oblog.showerr:Response.End()
isguest=1
password=Trim(Request.form("password"))
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
		if not oblog.codepass then oblog.ShowMsg "验证码错误，请重新输入","back"
	End If
end if
if checkuserlogined And password = "" Then username = oblog.l_uname:isguest=0
'Process...
set rs=oblog.execute("select userid,isencomment,ishide from oblog_album where fileid="&fileid)
if rs.eof then Response.Write("参数错误"):set rs=nothing:Response.End()
If rs("isencomment")<>"1" Or rs("ishide") = 1 Then Response.Write("该相片不允许回复"):set rs=nothing:Response.End()
mainuserid=rs(0)
set rs=Server.CreateObject("adodb.recordset")
rs.open "select top 1 * from oblog_albumcomment",conn,2,2
rs.addnew
rs("mainid")=fileid
rs("userid")=mainuserid
rs("comment_user")=EncodeJP(username)
rs("commenttopic")=EncodeJP(oblog.InterceptStr(oblog.filt_badword(commenttopic),250))
rs("comment")=comment
rs("homepage")=oblog.InterceptStr(oblog.filt_badword(Request.Form("homepage")),250)
rs("addtime")=oblog.ServerDate(now())
rs("addip")=oblog.userip
rs("isguest")=isguest
rs("istate")=oblog.CacheConfig(50)
rs.update
rs.close
set rs=Nothing
Dim scores
If oblog.CacheConfig(50) = 0 Then
	scores=0
Else
	scores=oblog.CacheScores(6)
End if
oblog.execute("update oblog_user set lastcomment='" & oblog.ServerDate(Now()) &"' where userid="&mainuserid)
Response.Cookies(cookies_name)("LastComment") = oblog.ServerDate(Now())
If oblog.CacheConfig(50)=0 Then
	oblog.ShowMsg "评论成功，审核后可见", ""
	Response.End
Else
	oblog.execute("update oblog_album set commentnum=commentnum+1 where fileid="&fileid)
	oblog.execute("update oblog_user set comment_count=comment_count+1,scores=scores+" & scores&" where userid="&mainuserid)
	oblog.execute("update oblog_setup set comment_count=comment_count+1")
End If
oblog.ShowMsg "评论成功", ""
%>