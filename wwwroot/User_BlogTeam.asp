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
		oblog.adderrstr("�û����Ʋ���Ϊ�գ�")
		oblog.showusererr
		exit sub
	End If
	If otheruser=oblog.l_uname Then
		oblog.adderrstr("�����Լ���Ϊ���ö���")
		oblog.showusererr
		exit sub
	End If
	set rs=oblog.execute("select en_blogteam,userid from [oblog_user] where username='"&otheruser&"'")
	if rs.eof then
		set rs=nothing
		oblog.adderrstr("�޴��û���")
		oblog.showusererr
		exit sub
	else
		if rs(0)<>1 then
			set rs=nothing
			oblog.adderrstr("���û�����Ϊ�����������Ŷӣ�")
			oblog.showusererr
			exit sub
		Else
			Set trs = oblog.Execute ("select * From oblog_blogteam WHERE otheruserid = " & rs(1)& " And mainuserid=" & oblog.l_uid)

			If Not trs.EOF Then
				oblog.adderrstr("���û��Ѿ�Ϊ�Ŷӳ�Ա��")
				oblog.showusererr
				exit Sub
			End if
			oblog.execute("insert into [oblog_blogteam] (mainuserid,otheruserid) values ("&oblog.l_uid&","&rs(1)&")")
			set rs=Nothing
			'��Ŀ���Ա����һ������Ϣ
			set rs=Server.CreateObject("adodb.recordset")
			rs.open "select top 1 * from oblog_pm Where 1=0",conn,1,3
			rs.addnew
			rs("incept")=oblog.Interceptstr(otheruser,50)
			rs("topic")="ϵͳ��Ϣ:���յ���ͬ׫д����"
			rs("content")="ϵͳ��Ϣ:" & vbcrlf & "[" & oblog.l_uname & " ]�������������Ĺ�ͬ׫д�ƻ����������ڷ�����־ʱ�ĸ߼�ѡ�����������־����������־�У����빲ͬ׫д��" & vbcrlf & "�����������룬������[��������]=>[��ͬ׫д]���˳���"
			rs("sender")=oblog.l_uname
			rs.update
			oblog.ShowMsg "�����Ŷӳ�Ա�ɹ�,�Ѿ���ó�Ա���Ͷ���֪ͨ!","user_blogteam.asp"
		end if
	end if
end sub

sub delotheruser
	if id="" then
		oblog.adderrstr("��ָ��ɾ��������")
		oblog.showusererr
		exit sub
	end if
	id=CLng(id)
	oblog.execute("delete  from [oblog_blogteam] where id="&id&" and ( mainuserid="&oblog.l_uid&" or otheruserid="&oblog.l_uid&" )")
	oblog.ShowMsg "�����ɹ���",""
end Sub

sub invite()
%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUp" class="FieldsetForm">
						<legend>�������Ѽ����ҵ��ŶӲ��ͣ�</legend>
						<form name="form1" method="post" action="user_blogteam.asp?action=addotheruser">
							<ul>
								<li>�������ѵ��ҵĲ����Ŷ�,�����ú��Ѱ���־�����ҵ�blog�����ھ������!��</li>
								<li><label>�û�����<input name="otheruser" type="text" id="otheruser" maxlength="20" /></label></li>
								<li><input type="submit" id="Submit" value=" �� �� " /></li>
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
						<td class="t2">�Ŷӳ�Ա����</td>
						<td class="t3">����</td>
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
								<a href="user_blogteam.asp?action=del&id=<%=rs(1)%>" <%="onClick='return confirm(""ɾ���󣬴˳�Ա������־�����������blog����ʾ,ȷ��Ҫɾ����"");'"%>><span class="red">ɾ���˲��ͳ�Ա</span></a>
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
						<td class="t2">�Ҽ�����Ŷ�</td>
						<td class="t3">����Ա</td>
						<td class="t4">����</td>
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
								<a href="user_blogteam.asp?action=del&id=<%=rs("id")%>" <%="onClick='return confirm(""ȷ���˳���"");'"%>><span class="red">�˳��˲����Ŷ�</span></a>
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