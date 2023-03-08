<!--#include file="user_top.asp"-->
<script language="javascript" src="inc/function.js"></script>
<%
dim rs,sql,blog,iCount,sGuide
dim id,cmd,action
Dim groupName,tsql
cmd=trim(request("cmd"))
action=trim(request("action"))
id=oblog.filt_badstr(trim(Request("id")))
if cmd="" then
	cmd=0
else
	cmd=clng(cmd)
end If
Set rs =oblog.execute ("SELECT g_name FROM oblog_groups WHERE groupid = " & oblog.l_uGroupId)
groupName = rs (0)
rs.close
G_P_FileName="user_pmmanage.asp?cmd=" & cmd & "&page="

tsql="or groups like '"&oblog.l_uGroupId&",%' or groups like '%,"&oblog.l_uGroupId&"' or groups like '%,"&oblog.l_uGroupId&",%' or groups ='"&oblog.l_uGroupId&"'"

select case action
	case "add"
	call add()
	case "saveadd"
	call saveadd()
	case "del"
	call del()
	case else
	call main()
end select
set rs=nothing
set blog=nothing
%>
</table>
</body>
</html>
<%
sub main()
%>

<%
	dim ssql,iCount,i,lPage,lAll,lPages,iPage,freen
	ssql="id,sender,incept,topic,addtime,isguest,isreaded,issys,content"
	select case cmd
		case 0
			sql="select "&ssql&" from oblog_pm where incept='"&oblog.l_uname&"' and delr=0 "&tsql&" order by issys desc,id desc"
			sGuide=sGuide & "收件箱"
		case 1
			sql="select "&ssql&" from oblog_pm where sender='"&oblog.l_uname&"' and dels=0 order by id desc"
			sGuide=sGuide & "发件箱"
		case else
	end select
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	rs.Open sql,Conn,1,1
	iCount=rs.RecordCount
  '分页处理
  lAll=INT(rs.recordcount)
    If lAll=0 Then
    	rs.Close
    	Set rs=Nothing
    	%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="user_pmmanage.asp">收件箱</a></li>
					<li><a href="user_pmmanage.asp?cmd=1">发件箱</a></li>
					<li><a href="javascript:openScript('user_pm.asp?action=send',450,400)">发送短信</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<!-- 没有相关记录 -->
					<div class="msg"><%=sGuide & " 没有相关纪录" %></div>
					<!-- 没有相关记录 end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/42.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
		<script>parent.show_title("<%=sGuide%>");</script>
    	<%
    	Exit Sub
    End If
    i=0
    iPage=20
	'分页
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = Int(Request("page"))
	End If

	'设置缓存大小 = 每页需显示的记录数目
	rs.CacheSize = iPage
	rs.PageSize = iPage
	rs.movefirst
	lPages = rs.PageCount
	If lPage>lPages Then lPage=lPages
	rs.AbsolutePage = lPage
	i=0

	freen=oblog.l_Group(27,0)-lAll
	%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="chk_idAll(myform,1);">全部选择</a></li>
					<li><a href="#" onclick="chk_idAll(myform,0);">全部取消</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'删除选中的短信吗?')==true) { document.myform.submit();}">删除短信</a></li>
					<li><a href="user_pmmanage.asp">收件箱</a></li>
					<li><a href="user_pmmanage.asp?cmd=1">发件箱</a></li>
					<li><a href="javascript:openScript('user_pm.asp?action=send',450,400)">发送短信</a></li>
					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="PmManageTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2">标题</td>
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
					<form name="myform" method="Post" action="user_pmmanage.asp?action=del&cmd=<%=cmd%>" onSubmit="return confirm('确定要执行选定的操作吗？');">
					<table id="PmManage" class="TableList" cellpadding="0">
						<%
						'Do while not rs.EOF
						Do While Not rs.Eof And i < rs.PageSize
						i = i + 1
						%>
						<tr id="u<%=rs("id")%>" onclick="chk_iddiv('<%=rs("id")%>')">
							<td class="t1" title="点击选中">
								<input onclick="chk_iddiv('<%=rs("id")%>')" name='id' type='checkbox' id="c<%=cstr(rs("id"))%>" value='<%=cstr(rs("id"))%>'<%If rs("issys")= 1 Then%>disabled<%End if%> />

							</td>
							<td class="t2">
								<%
								If rs("issys")= 1 Then
									response.write "<span class=""red"">系统通知</span>"
								Else
									response.write OB_IIF2(rs("isreaded"),"<span class=""grey"">已读</span>","<span class=""red"">未读</span>")
								End if
								%>
								<a href="javascript:openScript('user_pm.asp?action=read<%=cmd%>&id=<%=rs("id")%>',450,380)" title="cssbody=[dvbdy1] cssheader=[dvhdr1] body=[<%=oblog.filt_html(rs("content"))%>]"><%=oblog.filt_html(rs("topic"))%></a><br />
								<span class="message_user">
									<%If cmd=1 Then%>
										To <span class="green"><%
												If rs("incept")="0" Then
													response.Write "<span style=""color:#090;font-weight:600;"">" &groupName& "</span>"
												Else
													response.Write oblog.filt_html(rs("incept"))
												End If
										%></span>
									<%Else%>
										From <span class="green"><%
												If rs("issys")= 1 Then
													Response.Write "" & rs("sender") &""
												Else
													Response.Write oblog.filt_html(rs("sender"))
												End If
										%></span>
									<%End If%>
								</span>
								<!--时间-->
								<div class="time">on&nbsp;<%=OB_IIF(rs("addtime"),"-")%></div>
							</td>
							<td class="t3">
									<%If cmd=0 Then%>
										<%If rs("issys")= 1 Then %>
										<%Else%>
											<a href="javascript:openScript('user_pm.asp?action=send&incept=<%=rs("sender")%>&topic=<%="回复:"&oblog.filt_html(rs("topic"))%>',450,400)"  title=""><span class="blue">回复</span></a>
											<a href = "user_pmmanage.asp?action=del&id=<%=rs("id")%>" onclick="return confirm('确定要删除吗？');"><span class="red">删除</span></a>
										<%End if%>
									<%Else%>
										<%If rs("issys")= 1 Then %>
										<%Else%>
											<a href = "user_pmmanage.asp?action=del&id=<%=rs("id")%>" onclick="return confirm('确定要删除吗？');"><span class="red">删除</span></a>
										<%End if%>
									<%End if%>
							</td>
						</tr>
						<%
							'If i>iPage Then Exit Do
							'rs.Movenext
							'Loop
							'rs.Close
							'Set rs = Nothing
						%>
						<%
							if i>=iPage then exit do
							rs.movenext
							loop
						%>
					</table>
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
	<script>parent.show_title("<%=sGuide%>");</script>
  <%
	rs.Close
	if iCount>oblog.l_Group(27,0) then
			'oblog.execute("update oblog_pm set delr=1 where id not in (select top "&oblog.l_gPmMax&" id from  oblog_pm where incept='"&oblog.l_uname&"' order by id desc ) and incept='"&oblog.l_uname&"'")
		oblog.execute("delete from oblog_pm where delr=1 and dels=1")
		Response.Write"<script language=JavaScript>alert(""您的信箱已满，请及时清理！"");</script>"
	end if
	set rs=Nothing
end sub


sub del()
	if id="" then
		oblog.adderrstr( "错误：请指定要删除的对象！")
		oblog.showusererr
		exit sub
	end if
	if instr(id,",")>0 then
		dim n,i
		id=FilterIDs(id)
		n=split(id,",")
		for i=0 to ubound(n)
			delone(n(i))
		next
	else
		delone(id)
	end if
	set rs=nothing
	oblog.ShowMsg "删除短消息成功！",""
end sub

sub delone(id)
	id=clng(id)
	select case cmd
		case 0
		sql="update oblog_pm set delr=1 where id=" & clng(id)&" and incept='"&oblog.l_uname&"'"
		case 1
		sql="update oblog_pm set dels=1 where id=" & clng(id)&" and sender='"&oblog.l_uname&"'"
	end select
	oblog.execute(sql)
end sub
%>