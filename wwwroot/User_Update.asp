<!--#include file="user_top.asp"-->
<%
dim action,blog
action=Request("action")
set blog=new class_blog
Server.ScriptTimeOut=999999999
select Case action
	Case "update_message"
		Call update_message
	Case "update_blog"
		Call update_blog(0)
	Case "update_blog1"
		Call update_blog(1)
	Case "update_alllog"
		Call update_alllog
	Case else
		Call main()
End Select
If action <> "" Then Session ("CheckUserLogined_"&oblog.l_uName) = ""
Oblog.CheckUserLogined()
set blog=Nothing
Set oblog = Nothing
%>
	</body>
</html>
<%
sub main
%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="Photo" class="FieldsetForm">
						<legend>发布更新静态页面：</legend>
						<form action="user_photo.asp?action=savemodify" method="post" name="oblogform">
							<ul>
								<li><input type="button" id="Submit" value="更新首页" onClick="window.location='user_update.asp?action=update_blog'" /></li>
								<li><span class="grey1">更新博客的首页，通常是遇到博客首页无法打开，评论、留言等标题不准确等情况进行此操作。</span></li>
								<li><input type="button" id="Submit" value="更新留言"  onclick="window.location='user_update.asp?action=update_message'" /></li>
								<li>更新留言板,新增、删除留言后进行此操作。<strong>（一般不用进行此操作，系统会定时自动更新）</strong></li>
								<li><input type="button" id="Submit" value="更新数据" onClick="window.location='user_update.asp?action=update_blog1'" /></li>
								<li><span class="grey1">更新统计数据，积分、访问量、日志数、评论数、等数据不准确或出现负数请进行此操作。</span></li>
								<li><input type="button" id="Submit" value="重新发布" onClick="window.location='user_update.asp?action=update_alllog'" /></li>
								<li><span class="grey1">重新发布博客全站，当遇到无法解决的错误请进行此操作，如果错误无法解决请尽快联系网站管理员。</span><br /><strong class="red"><%If oblog.l_Group(32,0)>0 Then %>
	每天只能进行该操作<%=oblog.l_Group(32,0)%>次！请谨慎使用！<%End If%></strong></li>

							</ul>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
	</table>
<%end sub

sub update_message()
	const p=4
	Response.Write("") & vbcrlf
	Response.Write("<div id=""prompt"">") & vbcrlf
	Response.Write("	<ul>") & vbcrlf
	blog.progress_init
	blog.progress Int(1/p*100),"读取用户数据..."
	blog.userid=oblog.l_uid
	blog.progress Int(2/p*100),"更新留言板..."
	blog.update_message 1
	blog.progress Int(3/p*100),"更新最新留言..."
	blog.update_newmessage oblog.l_uid
	blog.progress Int(4/p*100),"更新留言完成!"
	Response.Write("		<li><a href='javascript:history.go(-1)'>返回上一页</a></li>") & vbcrlf
	Response.Write("	</ul>") & vbcrlf
	Response.Write("</div>") & vbcrlf
end sub

sub update_blog(itype)
	dim rsu,rst,n_log,n_comment,n_message,user_upfiles_num
	Dim RSDigg,DiggNum
	const p=16
	'On Error Resume Next 
	Response.Write("") & vbcrlf
	Response.Write("<div id=""prompt"">") & vbcrlf
	Response.Write("	<ul>") & vbcrlf
	blog.progress_init
	blog.progress Int(1/p*100),"更新全站统计数据..."
	set rsu=oblog.execute("select count(logid) from oblog_log where isdel=0 and isdraft=0 and userid="&oblog.l_uid)
	if not rsu.eof then n_log=rsu(0) else n_log=0
	set rsu=oblog.execute("select count(commentid) from oblog_comment where isdel=0 AND  istate=1 and userid="&oblog.l_uid)
	if not rsu.eof then n_comment=rsu(0) else n_comment=0
	set rsu=oblog.execute("select count(commentid) from oblog_albumcomment where isdel=0 AND  istate=1 and userid="&oblog.l_uid)
	if not rsu.eof then n_comment=n_comment + rsu(0)
	set rsu=oblog.execute("select count(messageid) from oblog_message where isdel=0 AND istate=1 and userid="&oblog.l_uid)
	if not rsu.eof then n_message=rsu(0) else n_message=0
	oblog.execute("update oblog_user set log_count="&n_log&",comment_count="&n_comment&",message_count="&n_message&" where userid="&oblog.l_uid)

	blog.progress Int(2/p*100),"更新分类统计数据..."
	set rst=Server.CreateObject("adodb.recordset")
	rst.open "select subjectid,subjectlognum,subjecttype from oblog_subject where userid="&oblog.l_uid,conn,2,2
	while not rst.eof
		If rst("subjecttype") = 0 Then
			set rsu=oblog.execute("select count(logid) from oblog_log where isdel=0 AND isdraft=0 AND   subjectid="&rst("subjectid"))
			if not rsu.eof then rst("subjectlognum")=rsu(0) else rst("subjectlognum")=0
		Else
			Set rsu = oblog.Execute ("SELECT COUNT(photoid) FROM oblog_album WHERE (ishide = 0 OR ishide IS NULL)  AND  userclassid = "&rst(0))
			if not rsu.eof then rst("subjectlognum")=rsu(0) else rst("subjectlognum")=0
		End if
		rst.update
		rst.movenext
	wend
	rst.close
	blog.progress Int(2.5/p*100),"更新用户推荐日志统计数据..."
	set RSDigg=Server.CreateObject("adodb.recordset")
	RSDigg.open "SELECT DiggID,DiggNum,Logid FROM oblog_userdigg WHERE authorid="&oblog.l_uid,CONN,2,2
	While Not RSDigg.EOF
		Set rsu = oblog.Execute ("SELECT COUNT(DID) FROM oblog_digg WHERE DiggID = "&RSDigg(0)&" AND authorid="&oblog.l_uid)
		If Not rsu.Eof Then  DiggNum = rsu(0) Else DiggNum = 0
		RSDigg(1) = DiggNum
		RSDigg.Update
		oblog.Execute ("UPDATE oblog_log SET DiggNum = "&DiggNum&" WHERE logid = "&RSDigg(2))
		RSDigg.MoveNext
	Wend
	set rst=nothing
	set rsu=Nothing
	Set RSDigg = Nothing

	blog.progress Int(3/p*100),"读取用户数据..."
	blog.userid=oblog.l_uid
	blog.progress Int(4/p*100),"更新首页..."
	blog.update_index 1
	blog.progress Int(5/p*100),"更新站点信息文件..."
	blog.update_info oblog.l_uid
	blog.progress Int(6/p*100),"生成新日志列表文件..."
	blog.update_newblog(oblog.l_uid)
	blog.progress Int(7/p*100),"更新最新留言..."
	blog.update_newmessage oblog.l_uid
	blog.progress Int(8/p*100),"生成首页日志分类文件..."
	blog.update_subject(oblog.l_uid)
	blog.progress Int(9/p*100),"生成功能页面..."
	blog.CreateFunctionPage
	blog.progress Int(10/p*100),"生成群组页面..."
	blog.update_mygroups(oblog.l_uid)
	blog.progress Int(11/p*100),"生成好友页面..."
	blog.update_friends(oblog.l_uid)
	blog.progress Int(12/p*100),"生成评论页面..."
	blog.update_comment(oblog.l_uid)
	blog.progress Int(13/p*100),"更新上传文件总数..."
	If Is_Sqldata = 1 Then
		oblog.execute ("UPDATE oblog_user SET user_upfiles_num = (select count(*) FROM oblog_upfile WHERE userid="&oblog.l_uid & ") WHERE userid="&oblog.l_uid)
	Else
		Set rsu = oblog.execute ("select count(*) FROM oblog_upfile WHERE userid="&oblog.l_uid )
		user_upfiles_num=RSU(0)
		rsu.close
		oblog.execute ("UPDATE oblog_user set user_upfiles_num = " &user_upfiles_num & " WHERE userid="&oblog.l_uid )
	End If
	blog.progress Int(14/p*100),"更新博客名..."
	'blog.update_blogname
	blog.progress Int(15/p*100),"更新公告..."
	'blog.update_placard (oblog.l_uid)
	if itype="0" Then
		blog.progress Int(16/p*100),"首页更新完成!"
	Else
		blog.progress Int(16/p*100),"更新统计数据完成!"
	End If
	Response.Write("		<li><a href='javascript:history.go(-1)'>返回上一页</a></li>") & vbcrlf
	Response.Write("	</ul>") & vbcrlf
	Response.Write("</div>") & vbcrlf
end sub

sub update_alllog()
	Dim updateblognum,lastlogid
	Dim trs
	If Int(oblog.l_Group(32,0)) = 1 Then
		If Not IsObject(conn) Then link_database
		Set trs = Server.CreateObject("adodb.recordset")
		trs.open "SELECT updateblognum,updateblogDate FROM oblog_user WHERE userid = "&oblog.l_uid, conn, 1, 3
		If IsNull(trs(1)) Or DateDiff("d",trs(1),Now()) > 0 Then
			updateblognum = 0
			trs(0) = 0
			trs(1) =Date()
			trs.Update
		Else
			updateblognum = ob_IIF(trs(0),0)
		End If
	End if
	lastlogid=Trim(Request("lastlogid"))
	if lastlogid<>"" then lastlogid=CLng(lastlogid) Else lastlogid=0
	If Int(oblog.l_Group(32,0)) = 1 Then
		if CLng(updateblognum)>=Int(oblog.l_Group(32,0)) and lastlogid=0 then
			oblog.adderrstr("整站更新每天只能进行"&oblog.l_Group(32,0)&"次!")
			oblog.showusererr
			exit sub
		end if
	End if
	if lastlogid=0 And Int(oblog.l_Group(32,0)) = 1  then
		updateblognum = updateblognum + 1
		trs(0) = updateblognum
		trs.Update
		trs.Close
		Set trs = Nothing
	end If
	Response.Write("") & vbcrlf
	Response.Write("<div id=""prompt"">") & vbcrlf
	Response.Write("	<ul>") & vbcrlf
	blog.progress_init
	blog.update_alllog oblog.l_uid,lastlogid
	Response.Write("		<li><a href='user_update.asp'>返回</a></li>") & vbcrlf
	Response.Write("	</ul>") & vbcrlf
	Response.Write("</div>") & vbcrlf
end sub
%>