<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/inc_ubb.asp"-->
<!--#include file="inc/inc_antispam.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/Class_qq.asp"-->
<%
Dim QueryString,i
Dim GroupId,postid,oTeam,cmd
QueryString = Request.QueryString
If InStr(QueryString,"404") > 0 Then
	QueryString = Right (QueryString,Len(QueryString)-InstrRev(QueryString,"?"))
	QueryString = Split(QueryString , "&")
	For i = 0 To UBound(QueryString)
		If InStr(QueryString(i),"cmd")>0 Then
			cmd = Replace(QueryString(i),"cmd=","")
		End If
		If InStr(QueryString(i),"gid")>0 Then
			GroupId = Replace(QueryString(i),"gid=","")
		End If
		If InStr(QueryString(i),"pid")>0 Then
			postid = Replace(QueryString(i),"pid=","")
		End if
	Next
Else
	GroupId=CLng(Request("gid")) '必须
	postid=Request("pid")
	cmd=Request("cmd")
End If
If postid<>"" Then postid=CLng(postid)
If postid<>""  Then
	If cmd="" Then cmd="show"
End If
Set oTeam=New Class_Team
oTeam.GroupId=GroupId
select Case cmd
	Case "show"
		'显示单篇内容
		oTeam.ShowPost postid
	Case "save"
		Call  oTeam.CheckQQLogin()
		'保存回复内容,返回该日志的最后一页
		Call oTeam.SaveComment()
	Case "join"
		'显示加入申请表单
		Call oTeam.ShowJoinForm()
	Case "pass"
		'处理申请:通过/拒绝[处理完成后给目标用户发短信息]
	Case "invite"
		'显示邀请表单,此处由会员在后台进行操作
	Case "links"
		Call oTeam.ShowlinksForm()
	Case "savelinks"
		Call oTeam.SaveLinks
	Case "placard"
		Call oTeam.ShowPlacardForm()
	Case "saveplacard"
		Call oTeam.Saveplacard
	Case "post"
		oTeam.PostForm
	Case "good0","good1","top0","top1","del"
		If postid<>"" Then Call oTeam.PostManage(cmd,postid)
	Case "users"
		Call oTeam.ShowUsers
	Case "wusers"
	Case "good"
		oTeam.ShowList(1)
	Case "savejoin"
		Call  oTeam.CheckQQLogin()
		oTeam.ActionJoin
	Case "postphoto"
		If oblog.CacheConfig(76) = "0" Then
			oblog.adderrstr("此功能已被系统关闭！")
			oblog.showerr
		End if
		Call oTeam.PostPHOTO
	Case "album"
		If oblog.CacheConfig(76) = "0" Then
			oblog.adderrstr("此功能已被系统关闭！")
			oblog.showerr
		End if
		Call oTeam.album()
	Case "photocomment"
		If oblog.CacheConfig(76) = "0" Then
			oblog.adderrstr("此功能已被系统关闭！")
			oblog.showerr
		End if
		Call oTeam.photocomment()
	Case "list"
		oTeam.ShowList(-1)
	Case Else
		oTeam.ShowList(0)
End select
Set oTeam=Nothing
%>