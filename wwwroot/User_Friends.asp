<!--#include file="user_top.asp"-->
<script language="javascript" src="inc/function.js"></script>
<%
Const PM_TITLE = "��ã����Ѿ�����Ϊ������"

Const SEND_PM = 1

Dim UserUrl,addFriendMsg
If oblog.CacheConfig(5)=1 Then
	If Left(oblog.l_udomain,8)="http://." Or Trim(oblog.l_udomain)="." Then
		UserUrl="<a href="""&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext&""" target=""_blank"">��������ҵĲ���</a>"
	Else
		UserUrl="<a href=""http://"&oblog.l_udomain&""" target=""_blank"">"&oblog.l_udomain&"</a>"
	End If
Else
	UserUrl="<a href="""&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext&""" target=""_blank"">��������ҵĲ���</a>"
End If
If true_domain=1 and oblog.l_ucustomdomain<>"" then
	UserUrl="<a href=""http://"&oblog.l_ucustomdomain&""" target=""_blank"">��������ҵĲ���</a>"
End If
addFriendMsg = " < a href = ""user_friends.asp?action=add"" > �������Ϊ���� </a> "

addFriendMsg = "��ӭ�����������ҵĲ���Ŷ��"

dim rs,sql,blog
dim id,cmd,action
cmd=Trim(Request("cmd"))
action=Trim(Request("action"))
id=Trim(Request("id") )
If id<>"" Then
	If Instr(id,",")>0 Then
		id=FilterIds(id)
	Else
		id=CLng(id)
	End If
End If
if cmd="" then
	cmd=0
else
	cmd=CLng(cmd)
end if
G_P_FileName="user_friends.asp?cmd=" & cmd & "&page="
if Request("page")<>"" then
    currentPage=cint(Request("page"))
else
	currentPage=1
end if
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

</body>
</html>
<%
sub main()
	dim ssql,i,lPage,lAll,lPages,iPage,sGuide,iCount
	sGuide=""
	ssql="id,username,nickname,user_icon1,blogname,user_dir,oblog_user.userid,oblog_friend.addtime,user_folder"
	select case cmd
		case 0
			sql="select "&ssql&" from oblog_friend,oblog_user where isblack=0 and oblog_friend.userid="&oblog.l_uid&" and oblog_friend.friendid=oblog_user.userid order by id desc"
			sGuide=sGuide & "�����б�"
		case 1
			sql="select "&ssql&" from oblog_friend,oblog_user where isblack=1 and oblog_friend.userid="&oblog.l_uid&" and oblog_friend.friendid=oblog_user.userid order by id desc"
			sGuide=sGuide & "������"
		case else
	end select
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'Response.Write(sql)
	rs.Open sql,Conn,1,3
	iCount=rs.RecordCount
	'��ҳ����
	lAll=Int(rs.recordcount)
    If lAll=0 Then
    	rs.Close
    	Set rs=Nothing
    	%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="purl('user_friends.asp','��������')">��������</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?action=add','��Ӻ���')">��Ӻ���</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?cmd=1','������')">������</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?action=add&type=black','��Ӻ���')">��Ӻ�����</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<!-- û����ؼ�¼ -->
					<div class="msg"><%=sGuide & " û����ؼ�¼" %></div>
					<!-- û����ؼ�¼ end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/42.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
    	<%
    	Exit Sub
    End If
    i=0
    iPage=12
	'��ҳ
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = Int(Request("page"))
	End If

	'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
	rs.CacheSize = iPage
	rs.PageSize = iPage
	rs.movefirst
	lPages = rs.PageCount
	If lPage>lPages Then lPage=lPages
	rs.AbsolutePage = lPage
	i=0
	Dim sKey
	If cmd="1" Then
		sKey="������"
	Else
		sKey="����"
	End If
%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="chk_idAll(myform,1);">ȫ��ѡ��</a></li>
					<li><a href="#" onclick="chk_idAll(myform,0);">ȫ��ȡ��</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ɾ��ѡ�е�������?')==true) { document.myform.submit();}">ɾ��</a></li>
					<li><a href="#" onclick="purl('user_friends.asp','��������')">��������</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?action=add','��Ӻ���')">��Ӻ���</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?cmd=1','������')">������</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?action=add&type=black','��Ӻ���')">��Ӻ�����</a></li>
					<li><a href="#" onclick="batchsend()">���Ͷ���</a></li>
					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="FriendsTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"></td>
						<td class="t3"><%=sKey%></td>
						<td class="t4">����</td>
						<td class="t5">����</td>
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
					<form name="myform" method="Post" action="user_friends.asp?action=del&cmd=<%=cmd%>" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
					<table id="Friends" class="TableList" cellpadding="0">
						<%
						do while not rs.eof
						%>
						<tr id="u<%=rs("id")%>"  onclick="chk_iddiv('<%=rs("id")%>')">
							<td class="t1" title="���ѡ��">
								<input name='id' type='checkbox' id="c<%=cstr(rs("id"))%>"   onclick="chk_iddiv('<%=cstr(rs("id"))%>')" value='<%=cstr(rs("id"))%>' />
							</td>
							<td class="t2">
								<a href="<%="go.asp?user="&rs("username")%>" target="_blank" class="user_icon"><img src="<%=ProIco(rs("user_icon1"),1)%>" /></a>
							</td>
							<td class="t3">
								<a href="<%="go.asp?user="&rs("username")%>" target="_blank"><%=OB_IIF(rs("nickname"),rs("username"))%></a>
								<!--ʱ��-->
								<div class="time">added&nbsp;on&nbsp;<%=rs("addtime")%></div>
							</td>
							<td class="t4">
								<a href="<%="go.asp?user="&rs("username")%>" target="_blank"><%=rs("blogname")%></a>
							</td>
							<td class="t5">
								<%
									Response.write " <a href=""javascript:openScript('user_pm.asp?action=send&incept="&oblog.filt_html(rs("username"))&"',450,400)""><span class=""blue"">������Ϣ</span></a>&nbsp;"
									Response.write "<a href='user_friends.asp?action=del&id=" & rs("id") &"' onClick='return confirm(""ȷ��Ҫɾ����"");'><span class=""red"">ɾ��</span></a>"
								%>
							</td>
						</tr>
						<%
							i=i+1
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
</table>
<%
end sub

sub add()
dim str1,str2
if Request("type")="black" then
	str1="��Ӻ�����"
	str2="�������û�����"
else
	str1="��Ӻ���"
	str2="�����û�����"
end if
%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="purl('user_friends.asp','��������')">��������</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?action=add','��Ӻ���')">��Ӻ���</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?cmd=1','������')">������</a></li>
					<li><a href="#" onclick="purl('user_friends.asp?action=add&type=black','��Ӻ���')">��Ӻ�����</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUp" class="FieldsetForm">
						<legend><%=str1%>��</legend>
						<form action="user_friends.asp?action=saveadd&type=<%=Request("type")%>" method="post" name="oblogform">
							<ul>
								<li><label><%=str2%><input name="friendname" type=text size="20" maxlength="30" value="<%=Request("friendname")%>" /></label></li>
							<%if Request("type")<>"black" then%>
								<li><label>ͬʱ���ģ�<input type="checkbox" value="true" name="is_sub" checked="checked" /></label></li>
							<%end if%>
								<li><input type="submit" name="addsubmit" id="Submit" value="<%=str1%>" onmouseup="window.close();" /></li>
								<li><span class="grey">
									<%if Request("type")="black" then%>
									����������Ժ󣬾Ͳ����ܵ�����û��Ķ���ɧ���ˡ�
									<%else%>
									���û���Ϊ���ѣ����Է���ķ���վ�ڶ��ţ������Ժͺ��ѹ���˽����־!
									<%end if%>
								</span></li>
							</ul>
							<input type="hidden" name="close" value="<%=Request("close")%>" />
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/42.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
end sub

sub saveadd()
	dim friendname,uid,isblack,surl,ublogname,rs1
	friendname=oblog.filt_badstr(Trim(Request("friendname")))
	if friendname="" then
		oblog.adderrstr("��������û�������Ϊ��")
		oblog.showusererr
		exit sub
	end if
	If friendname=oblog.l_uname Then
		oblog.adderrstr("���󣺲��ܽ��Լ����Ϊ�Լ��ĺ��ѻ������!")
		oblog.showusererr
		exit sub
	End If
	if Request("type")="black" then isblack=1 else isblack=0
	sql="select userid,blogname,user_dir,user_folder from oblog_user where username='"&friendname&"'"
	set rs=oblog.execute(sql)
	if rs.eof then
		oblog.adderrstr("�����޴��û�")
		oblog.showusererr
		exit sub
	end if
	uid=rs("userid")
	surl=blogurl&rs("user_dir")&"/"&rs("user_folder")&"/rss2.xml"
	ublogname=rs("blogname")
	set rs=oblog.execute("select id from oblog_friend where userid="&oblog.l_uid&" and friendid="&uid&" and isblack="&isblack)
	if rs.eof then
		oblog.execute("insert into [oblog_friend] (userid,friendid,isblack) values ("&oblog.l_uid&","&uid&","&isblack&")")
		If SEND_PM = 1 Then Call SendPM(uid)
		update_friends()
		'д�붩��
		if isblack=0 and Request("is_sub")="true" then
			Set rs1=Server.CreateObject("Adodb.Recordset")
			rs1.Open "select * From oblog_myurl Where userid="&oblog.l_uid&" and url='"&oblog.filt_badstr(surl)&"'",conn,1,3
			if not rs1.eof then
				rs1.close
			else
				rs1.AddNew
				rs1("classid") = 0
				rs1("subjectid")=0
				rs1("url")=sUrl
				rs1("userid")=oblog.l_uid
				rs1("addtime") = oblog.ServerDate(Now)
				rs1("encodeing")="gb2312"
				rs1("title")=ublogname
				rs1("mainuserid")=uid
				rs1("isupdate")=1
				rs1.Update
				rs1.Close
				set rs1=nothing
				oblog.execute("update oblog_user set sub_num=sub_num+1 where userid="&uid)
			end if
		end if
		set rs=Nothing
		If request("close") = "true" Then
			oblog.ShowMsg "��ӳɹ�","close"
		Else
			oblog.ShowMsg "��ӳɹ�","user_friends.asp"
		End if
	else
		set rs=nothing
		oblog.adderrstr("���󣺴��û��Ѿ����б���")
		oblog.showusererr
	end if
end sub


sub del()
	if id="" then
		oblog.adderrstr( "������ָ��Ҫɾ���Ķ���")
		oblog.showusererr
		exit sub
	end if
	If Instr(Id,",")>0 Then
		oblog.execute("delete from [oblog_friend] where userid=" & oblog.l_uid &" and id In ("&id & ")")
	Else
		oblog.execute("delete from [oblog_friend] where userid=" & oblog.l_uid &" and id ="&id )
	End If
	update_friends()
	oblog.ShowMsg "ɾ���ɹ���",""
end sub

sub update_friends()
	dim blog
	set blog=new class_blog
	blog.userid=oblog.l_uid
	blog.update_friends oblog.l_uid
	set blog=nothing
end Sub

Sub SendPM(ByVal inceptid )
	If SEND_PM <> 1 Then Exit Sub
	Dim sql,rs1,incept
	Set rs1 = oblog.Execute ("SELECT username FROM oblog_user WHERE userid = "&inceptid)
	If Not rs1.Eof Then
		incept =rs1(0)
	Else
		Exit Sub
	End if
	sql="select top 1 * from oblog_pm Where 1=0"
	set rs1=Server.CreateObject("adodb.recordset")
	rs1.open sql,conn,1,3
	rs1.addnew
	rs1("incept")= incept
	rs1("topic") = PM_TITLE
	rs1("content")=addFriendMsg
	rs1("sender")=oblog.l_uname
	rs1.update
End Sub
%>
<script language="javascript">
function batchsend(){
	var ids=read_checkbox('id');
	//alert(ids);
	openScript('user_pm.asp?action=batchsend&incept='+ids,450,400);
	}

</script>