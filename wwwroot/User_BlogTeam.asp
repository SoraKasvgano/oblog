<!--#include file="user_top.asp"-->
<%
dim action,id
action=Request.QueryString("action")
id=Trim(Request("id"))
Dim DivId
DivId=Request("div")
If DivId="" Then DivId=31
DivId=Cint(DivId)
select case action
	case "addotheruser"
		call addotheruser()
	case "del"
		call delotheruser()
	case Else
		If DivId = 31 Then
			call M_team()
		ElseIf DivId = 32 Then
			call J_team()
		ElseIf DivId = 33 Then
			Call Invite()
		End if
end select
%>
</body>
</html>
<%

sub addotheruser()
	dim otheruser,rs,trs
	otheruser=oblog.filt_badstr(Trim(Request.Form("otheruser")))
	If otheruser="" Then
		oblog.adderrstr("用户名称不能为空！")
		oblog.showusererr
		exit sub
	End If
	If otheruser=oblog.l_uname Then
		oblog.adderrstr("请勿将自己作为设置对象！")
		oblog.showusererr
		exit sub
	End If
	set rs=oblog.execute("select en_blogteam,userid from [oblog_user] where username='"&otheruser&"'")
	if rs.eof then
		set rs=nothing
		oblog.adderrstr("无此用户！")
		oblog.showusererr
		exit sub
	else
		if rs(0)<>1 then
			set rs=nothing
			oblog.adderrstr("该用户设置为不允许被加入团队！")
			oblog.showusererr
			exit sub
		Else
			Set trs = oblog.Execute ("select * From oblog_blogteam WHERE otheruserid = " & rs(1)& " And mainuserid=" & oblog.l_uid)

			If Not trs.EOF Then
				oblog.adderrstr("该用户已经为团队成员！")
				oblog.showusererr
				exit Sub
			End if
			oblog.execute("insert into [oblog_blogteam] (mainuserid,otheruserid) values ("&oblog.l_uid&","&rs(1)&")")
			set rs=Nothing
			'给目标成员发送一条短信息
			set rs=Server.CreateObject("adodb.recordset")
			rs.open "select top 1 * from oblog_pm Where 1=0",conn,1,3
			rs.addnew
			rs("incept")=oblog.Interceptstr(otheruser,50)
			rs("topic")="系统信息:您收到共同撰写邀请"
			rs("content")="系统信息:" & vbcrlf & "[" & oblog.l_uname & " ]邀请您参与他的共同撰写计划，您可以在发布日志时的高级选项里，将您的日志发布到其日志中，参与共同撰写。" & vbcrlf & "如果您不想参与，您可在[博客设置]=>[共同撰写]中退出。"
			rs("sender")=oblog.l_uname
			rs.update
			oblog.ShowMsg "加入团队成员成功,已经向该成员发送短信通知!","user_blogteam.asp"
		end if
	end if
end sub

sub delotheruser
	if id="" then
		oblog.adderrstr("请指定删除参数！")
		oblog.showusererr
		exit sub
	end if
	id=CLng(id)
	oblog.execute("delete  from [oblog_blogteam] where id="&id&" and ( mainuserid="&oblog.l_uid&" or otheruserid="&oblog.l_uid&" )")
	oblog.ShowMsg "操作成功！",""
end Sub

sub invite()
%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUp" class="FieldsetForm">
						<legend>邀请朋友加入我的团队博客：</legend>
						<form name="form1" method="post" action="user_blogteam.asp?action=addotheruser">
							<ul>
								<li>加入朋友到我的博客团队,可以让好友把日志发表到我的blog，现在就邀请吧!　</li>
								<li><label>用户名：<input name="otheruser" type="text" id="otheruser" maxlength="20" /></label></li>
								<li><input type="submit" id="Submit" value=" 邀 请 " /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%end sub%>
<%
Sub M_team()%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr2">
			<th>
				<table id="M_teamTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2">团队成员管理</td>
						<td class="t3">操作</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<table id="M_team" class="TableList" cellpadding="0">
<%
dim rs,i
set rs=oblog.execute("select a.username,b.id, a.user_dir,a.user_folder from oblog_user a,oblog_blogteam b where a.userid=b.otheruserid and b.mainuserid="&oblog.l_uid)
while not rs.eof
	i=i+1
%>
						<tr>
							<td class="t1">
								<%=i%>
							</td>
							<td class="t2">
								<a href="<%=rs("user_dir")&"/"&rs("user_folder")&"/index."&f_ext%>" target="_blank"><%=rs(0)%></a>
							</td>
							<td class="t3">
								<a href="user_blogteam.asp?action=del&id=<%=rs(1)%>" <%="onClick='return confirm(""删除后，此成员所有日志将不会在你的blog中显示,确定要删除吗？"");'"%>><span class="red">删除此博客成员</span></a>
							</td>
						</tr>
<%
	rs.movenext
wend
set rs=nothing
%>
					</table>
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/18.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
End Sub
Sub J_team()
%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr2">
			<th>
				<table id="J_teamTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2">我加入的团队</td>
						<td class="t3">管理员</td>
						<td class="t4">操作</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<table id="J_team" class="TableList" cellpadding="0">
<%
Dim rs,i
set rs=oblog.execute("select a.id,a.mainuserid,b.blogname,b.user_dir,b.username,b.user_folder from oblog_blogteam a,oblog_user b where a.otheruserid="&oblog.l_uid&" and b.userid=a.mainuserid")
i=0
while not rs.eof
	i=i+1
%>
						<tr>
							<td class="t1">
								<%=i%>
							</td>
							<td class="t2">
								<a href="<%=rs("user_dir")&"/"&rs("user_folder")&"/index."&f_ext%>" target="_blank"><%=rs(2)%></a>
							</td>
							<td class="t3">
								<a href="<%=rs("user_dir")&"/"&rs("user_folder")&"/index."&f_ext%>" target="_blank"><%=rs(4)%></a>
							</td>
							<td class="t4">
								<a href="user_blogteam.asp?action=del&id=<%=rs("id")%>" <%="onClick='return confirm(""确认退出吗？"");'"%>><span class="red">退出此博客团队</span></a>
							</td>
						</tr>
<%
rs.movenext
wend
%>
					</table>
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/18.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
End Sub
%>