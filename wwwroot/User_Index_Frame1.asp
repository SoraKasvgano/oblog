<!--#include file="user_top.asp"-->
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<table id="IndexFrame" class="TableList" cellpadding="0">
						<tr>
							<td class="t1">
								<div id="Welcome">
									<ul class="Login">
										<li><%=oblog.l_uname%>����ӭ����</li>
										<li>�ϴε�¼��<%=oblog.l_ulastlogin%></li>
										<li>��ǰʱ�䣺<%=now%></li>
										<%If oblog.l_blogpassword=1 Then %><li style="color:green;">���Ĳ���������ȫվ����״̬�������˽���������������Ĳ��ͣ�</li><%End If %>
									</ul>
									<ul class="NewInfo">
										<li>����һ����������&nbsp;<a href="#" onclick="purl('user_comments.asp','��־����')"><%=FmtMinutes(oblog.l_ulastcomment)%></a>&nbsp;ǰ</li>
										<li>����һ����������&nbsp;<a href="#" onclick="purl('user_messages.asp','�ÿ�����')"><%=FmtMinutes(oblog.l_ulastmessage)%></a>&nbsp;ǰ</li>
<%
	set rs=oblog.execute("select count(teamid) from oblog_teamusers where userid="&oblog.l_uid&" and state=1")
	if rs(0)>0 then
		Response.Write("										<li><img src=""oBlogStyle/UserAdmin/7/newmsg.gif"" align=""absmiddle"" />�յ� <a href='#' onclick=""purl('user_team.asp?action=members&cmd=2','�������')"">"&rs(0)&"</a> ��" &oblog.CacheConfig(69)& "�������</a></li>")
	end if
	set rs=oblog.execute("select count(userid) from oblog_teamusers where state=2 and teamid in (select a.teamid from oblog_team a,oblog_teamusers b where a.teamid=b.teamid and b.state=5 and  b.userid="&oblog.l_uid&")")
	if rs(0)>0 then
		Response.Write("										<li><img src=""oBlogStyle/UserAdmin/7/newmsg.gif"" align=""absmiddle"" />�յ� <a href='#' onclick=""purl('user_team.asp?action=members&cmd=4','�鿴����')"">"&rs(0)&"</a> ���û���" &oblog.CacheConfig(69)& "����</a></li>")
	end if
%>
									</ul>
									<div class="clear"></div>
								</div>
								<div id="InfoTemplate">
									<div class="Info">
										<ul class="UserInfo">
											<li class="l1"><img class="face" src="<%=oblog.l_uIco%>" align="absmiddle" /></li>
											<li class="l2">�ǳƣ�<%=oblog.l_unickname%></li>
											<li class="l3">����<%=oblog.l_Group(1,0)%></li>
											<li class="l4">���֣�<%=oblog.l_uScores%></li>
											<li class="l5"><input type="button" value="�޸��ҵ����ϼ�ͷ��" onclick="purl('user_setting.asp?action=userinfo&div=13','��������')"></li>
											<li class="l6"><input type="button" value="�ʺŰ�ȫ����" onclick="purl('user_setting.asp?action=userpassword&div=12','��������')"></li>
										</ul>
										<ul class="BlogInfo">
											<li class="l1">��־����:<%=oblog.l_ulogcount%></li>
											<li class="l2">��������:<%=oblog.l_ucommentcount%></li>
											<li class="l3">��������:<%=oblog.l_umessagecount%></li>
											<li class="l4">���ʴ���:<%=oblog.l_uvisitcount%></li>
										</ul>
										<%
										'�������ݼ���
										'l_gUpSpace=0 ����/-1�������ϴ�
										Dim freesize, maxsize,maxsize1,thisPercent
										maxsize1 = oblog.l_Group(24,0)
										If maxsize1>0 Then
											maxsize = oblog.showsize(maxsize1 * 1024)
											freesize = oblog.showsize(Int(maxsize1*1024 - oblog.l_uUpUsed))
											thisPercent=oblog.l_uUpUsed/(maxsize1*1024)*100
										Elseif maxsize1=0 Then
											maxsize = "����"
											freesize = "����"
											thisPercent=0
										Elseif maxsize1=-1 Then
											maxsize = 0
											freesize = 0
											thisPercent=100
										End If
										%>
										<div id="space">
											<table cellpadding="0" title="ʹ�ÿռ䣺<%=oblog.showsize(oblog.l_uUpUsed)%>
					ʣ��ռ䣺<%=freesize%>">
												<tr>
													<td class="used" width="<%=thispercent%>%" height="12"></td>
													<td width="100%"></td>
												</tr>
											</table>
											<ul>
												<li class="l1">ʹ�ÿռ䣺<span class="red"><%=oblog.showsize(oblog.l_uUpUsed)%></span></li>
												<li class="l2">ʣ��ռ䣺<span class="red"><%=freesize%></span></li>
												<li class="l3">�ռ��С��<span class="red"><%=maxsize%></span></li>
											</ul>
										</div>
									</div>
									<div class="Template">
									<%
										Dim rs
										Set rs=oblog.Execute("select top 2 id,userskinname,skinauthorurl,skinauthor,skinpic From oblog_userskin Where ispass=1 Order By id Desc")
										If Not rs.Eof Then
										Do While Not rs.Eof
									%>
										<ul>
											<li class="img"><a href="showskin.asp?id=<%=rs(0)%>" target="_blank" title="���Ԥ��"><img src="<%=ProIco(rs(4),3)%>" /></a></li>
											<li class="name"><a href="showskin.asp?id=<%=rs(0)%>"  target="_blank" title="���Ԥ��"><%=rs(1)%></a></li>
										</ul>
									<%
										rs.Movenext
										Loop
										End If
										Set rs=Nothing
									%>
									</div>
									<div class="clear"></div>
								</div>
							</td>
							<td class="t2">
								<div id="SitePlacard">
									<div class="top">վ�㹫��</div>
									<div class="content">
<%=oblog.setup(7,0)%>
									</div>
									<div class="clear"></div>
								</div>
								<!-- ��̨���λ �Ƽ���С313*100 -->
								<div id="AD">
									<div class="content" style="display:block;width:313px;height:100px;overflow:hidden;"><%
									On Error Resume Next 
									server.execute(oblog.CacheConfig(80)&"/gg_user_desktop_main.htm")
									If Err Then Err.clear:response.write "δ���ú�̨���λ!" 
									%></div>
									<div class="clear"></div>
								</div>
								<!-- ��̨���λ �Ƽ���С313*100 END -->
								<!-- ��ݷ�ʽ -->
								<div id="btu">
									<ul>
										<li class="l1"><a href="#" onclick="purl('user_friendurl.asp','��������')" title="���ò�����������">��������</a></li>
										<li class="l2"><a href="#" onclick="purl('user_setting.asp?action=placard&div=12','��������')" title="���ò��͹���">���͹���</a></li>
										<li class="l3"><a href="#" onclick="purl('user_setting.asp?action=blogpassword&div=15','��������')" title="��������������">���ܲ���</a></li>
										<li class="l4"><a href="#" onclick="purl('user_setting.asp?action=userinfo&div=21','��������')" title="���ø�������">��������</a></li>
										<li class="l5"><a href="#" onclick="purl('user_setting.asp?action=userpassword&div=23','��������')" title="�ʺ����뱣��">���뱣��</a></li>
										<li class="l6"><a href="#" onclick="purl('user_setting.asp?action=blogstar&div=16','��������')" title="���벩��֮��">���벩��</a></li>
									</ul>
								</div>
								<!-- ��ݷ�ʽ END -->
							</td>
						</tr>
					</table>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
</body>
</html>