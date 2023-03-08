<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="Inc/Class_TrackBack.asp" -->
<%
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
ON Error Resume Next
If Not lcase(Request.ServerVariables("REQUEST_METHOD"))="post" Then Response.End
If oblog.CacheConfig(54) = "1" Then Response.write("系统永久禁止引用通告功能!"):Response.End
If Application(cache_name_user&"_systemenmod")<>"" Then
	Dim enStr
	enStr=Application(cache_name_user&"_systemenmod")
	enStr=Split(enStr,",")
	If enStr(3)="1" Then	Response.write("系统临时禁止使用引用通告功能!"):Response.End
End if
Dim objTrackback
Dim LogId,IP,url,title,BlogName,Excerpt,rst,rstCache
'恶意内容的操作不返回XML返回值
'IP检测
oblog.chk_commenttime
'tb.asp?id=53&TBcode=200703210942u85Csg1RRm6O&url=http://lj/oblog41/go.asp&blog_name=atai&title=好文章&excerpt=不错啊
LogId=Request("id")
logId=CLng(LogId)
IP=GetIP
Url=Trim(Request("url"))
title=Trim(Request("title"))
BlogName=Trim(Request("blog_name"))
Excerpt=Trim(Request("excerpt"))
If url=blogdir&"tb.asp" Then
'如果url为空则停止相应
	Response.End
End if
'内容检测
if oblog.chk_badword(url)>0 then oblog.adderrstr("地址中含有系统不允许的字符！")
if oblog.chk_badword(title)>0 then oblog.adderrstr("标题中含有系统不允许的字符！")
if oblog.chk_badword(BlogName)>0 then oblog.adderrstr("BLOG名称中含有系统不允许的字符！")
if oblog.chk_badword(Excerpt)>0 then oblog.adderrstr("摘要中含有系统不允许的字符！")
'专属关键字判定
if oblog.errstr<>"" Then oblog.showerr(): Response.End()
Call Link_Database
'频度检测,如果同一IP在单位时间内发布的通告申请达到一定限额，则自动封IP
'If oblog.ChkWhiteIP(IP) = False Then
	Set rst=oblog.Execute("select count(id) From Oblog_trackback Where ip='" & IP & "' And datediff("&G_Sql_mi&",addtime,"&G_Sql_Now&")<="&oblog.CacheConfig(66))
	If rst(0)> Int(oblog.CacheConfig(65)) Then
		'加入黑名单
		oblog.KillIp(IP)
'		oblog.ShowMsg "因为您的一些操作对系统进行了干扰，你的IP被加入黑名单",""
		Response.End
	End If
	rst.Close
	Set rst=Nothing
'End if

'进行接收环节的处理
Set objTrackback = New Class_TrackBack
objTrackback.LOGID=LogId
objTrackback.IP=IP
objTrackback.URL=Url
'objTrackback.TBUSER=Trim(Request.QueryString("tbuser"))
objTrackback.TITLE=title
objTrackback.BLOG_NAME=BlogName
objTrackback.EXCERPT=Excerpt
Response.Cookies(cookies_name)("LastComment") = oblog.ServerDate(Now())
If objTrackback.CheckTB (LCase(Trim(Request("TBcode")))) Then Call objTrackback.Receive()
Set objTrackback=Nothing
conn.Close
Set conn=Nothing
%>