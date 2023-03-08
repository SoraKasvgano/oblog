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
										<li><%=oblog.l_uname%>，欢迎您！</li>
										<li>上次登录：<%=oblog.l_ulastlogin%></li>
										<li>当前时间：<%=now%></li>
										<%If oblog.l_blogpassword=1 Then %><li style="color:green;">您的博客现在是全站加密状态，其他人将不能正常浏览您的博客！</li><%End If %>
									</ul>
									<ul class="NewInfo">
										<li>最新一条评论是在&nbsp;<a href="#" onclick="purl('user_comments.asp','日志评论')"><%=FmtMinutes(oblog.l_ulastcomment)%></a>&nbsp;前</li>
										<li>最新一条留言是在&nbsp;<a href="#" onclick="purl('user_messages.asp','访客留言')"><%=FmtMinutes(oblog.l_ulastmessage)%></a>&nbsp;前</li>
<%
	set rs=oblog.execute("select count(teamid) from oblog_teamusers where userid="&oblog.l_uid&" and state=1")
	if rs(0)>0 then
		Response.Write("										<li><img src=""oBlogStyle/UserAdmin/7/newmsg.gif"" align=""absmiddle"" />收到 <a href='#' onclick=""purl('user_team.asp?action=members&cmd=2','邀请加入')"">"&rs(0)&"</a> 个" &oblog.CacheConfig(69)& "邀请加入</a></li>")
	end if
	set rs=oblog.execute("select count(userid) from oblog_teamusers where state=2 and teamid in (select a.teamid from oblog_team a,oblog_teamusers b where a.teamid=b.teamid and b.state=5 and  b.userid="&oblog.l_uid&")")
	if rs(0)>0 then
		Response.Write("										<li><img src=""oBlogStyle/UserAdmin/7/newmsg.gif"" align=""absmiddle"" />收到 <a href='#' onclick=""purl('user_team.asp?action=members&cmd=4','查看申请')"">"&rs(0)&"</a> 个用户的" &oblog.CacheConfig(69)& "申请</a></li>")
	end if
%>
									</ul>
									<div class="clear"></div>
								</div>
								<div id="InfoTemplate">
									<div class="Info">
										<ul class="UserInfo">
											<li class="l1"><img class="face" src="<%=oblog.l_uIco%>" align="absmiddle" /></li>
											<li class="l2">昵称：<%=oblog.l_unickname%></li>
											<li class="l3">级别：<%=oblog.l_Group(1,0)%></li>
											<li class="l4">积分：<%=oblog.l_uScores%></li>
											<li class="l5"><input type="button" value="修改我的资料及头像" onclick="purl('user_setting.asp?action=userinfo&div=13','博客设置')"></li>
											<li class="l6"><input type="button" value="帐号安全设置" onclick="purl('user_setting.asp?action=userpassword&div=12','博客设置')"></li>
										</ul>
										<ul class="BlogInfo">
											<li class="l1">日志总数:<%=oblog.l_ulogcount%></li>
											<li class="l2">评论数量:<%=oblog.l_ucommentcount%></li>
											<li class="l3">留言数量:<%=oblog.l_umessagecount%></li>
											<li class="l4">访问次数:<%=oblog.l_uvisitcount%></li>
										</ul>
										<%
										'进行数据计算
										'l_gUpSpace=0 限制/-1不允许上传
										Dim freesize, maxsize,maxsize1,thisPercent
										maxsize1 = oblog.l_Group(24,0)
										If maxsize1>0 Then
											maxsize = oblog.showsize(maxsize1 * 1024)
											freesize = oblog.showsize(Int(maxsize1*1024 - oblog.l_uUpUsed))
											thisPercent=oblog.l_uUpUsed/(maxsize1*1024)*100
										Elseif maxsize1=0 Then
											maxsize = "不限"
											freesize = "不限"
											thisPercent=0
										Elseif maxsize1=-1 Then
											maxsize = 0
											freesize = 0
											thisPercent=100
										End If
										%>
										<div id="space">
											<table cellpadding="0" title="使用空间：<%=oblog.showsize(oblog.l_uUpUsed)%>
					剩余空间：<%=freesize%>">
												<tr>
													<td class="used" width="<%=thispercent%>%" height="12"></td>
													<td width="100%"></td>
												</tr>
											</table>
											<ul>
												<li class="l1">使用空间：<span class="red"><%=oblog.showsize(oblog.l_uUpUsed)%></span></li>
												<li class="l2">剩余空间：<span class="red"><%=freesize%></span></li>
												<li class="l3">空间大小：<span class="red"><%=maxsize%></span></li>
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
											<li class="img"><a href="showskin.asp?id=<%=rs(0)%>" target="_blank" title="点击预览"><img src="<%=ProIco(rs(4),3)%>" /></a></li>
											<li class="name"><a href="showskin.asp?id=<%=rs(0)%>"  target="_blank" title="点击预览"><%=rs(1)%></a></li>
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
									<div class="top">站点公告</div>
									<div class="content">
<%=oblog.setup(7,0)%>
									</div>
									<div class="clear"></div>
								</div>
								<!-- 后台广告位 推荐大小313*100 -->
								<div id="AD">
									<div class="content" style="display:block;width:313px;height:100px;overflow:hidden;"><%
									On Error Resume Next 
									server.execute(oblog.CacheConfig(80)&"/gg_user_desktop_main.htm")
									If Err Then Err.clear:response.write "未设置后台广告位!" 
									%></div>
									<div class="clear"></div>
								</div>
								<!-- 后台广告位 推荐大小313*100 END -->
								<!-- 快捷方式 -->
								<div id="btu">
									<ul>
										<li class="l1"><a href="#" onclick="purl('user_friendurl.asp','博客设置')" title="设置博客友情连接">友情连接</a></li>
										<li class="l2"><a href="#" onclick="purl('user_setting.asp?action=placard&div=12','博客设置')" title="设置博客公告">博客公告</a></li>
										<li class="l3"><a href="#" onclick="purl('user_setting.asp?action=blogpassword&div=15','博客设置')" title="给博客设置密码">加密博客</a></li>
										<li class="l4"><a href="#" onclick="purl('user_setting.asp?action=userinfo&div=21','博客设置')" title="设置个人资料">个人资料</a></li>
										<li class="l5"><a href="#" onclick="purl('user_setting.asp?action=userpassword&div=23','博客设置')" title="帐号密码保护">密码保护</a></li>
										<li class="l6"><a href="#" onclick="purl('user_setting.asp?action=blogstar&div=16','博客设置')" title="申请博客之星">申请博星</a></li>
									</ul>
								</div>
								<!-- 快捷方式 END -->
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