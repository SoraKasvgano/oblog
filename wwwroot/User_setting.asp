<!--#include file="user_top.asp"-->
<%
Dim DivId
DivId=Request("div")
If DivId="" Then DivId=11
DivId=Cint(DivId)
%>
<script src="oBlogStyle/move.js" type="text/javascript"></script>
<script src="inc/function.js" type="text/javascript"></script>
<script type="text/javascript">
function getImg(){
	if (document.oblogform.ico.value!=""){
		document.oblogform.imgIcon.src=document.oblogform.ico.value;
	}
}
</script>
<table id="TableBody" class="Setting" cellpadding="0">
	<thead id="TableBody_thead">
		<tr>
			<th>
				<ul id="TabPage2">
					<li id="left_tab1" <%If divId=11 or divId=12 or divId=13 or divId=14 or divId=15 or divId=16 Then%>class="Selected"<%End If%> onClick="javascript:border_left('TabPage2','left_tab1');self.location='user_setting.asp?action=0&div=11'" title="博客设置">博客设置</li>
					<li id="left_tab2" <%If divId=21 or divId=22 or divId=23 Then%>class="Selected"<%End If%> onClick="javascript:border_left('TabPage2','left_tab2');self.location='user_setting.asp?action=userinfo&div=21'" title="博客设置">用户设置</li>
					<li id="left_tab3" <%If divId=31 or divId=32 or divId=33 Then%>class="Selected"<%End If%> onClick="javascript:border_left('TabPage2','left_tab3');self.location='user_setting.asp?action=blogteam&div=31'" title="博客设置">共同撰写</li>
				</ul>

				<div id="left_menu_cnt">
					<ul id="dleft_tab1" <%If divId=11 or divId=12 or divId=13 or divId=14 or divId=15 or divId=16 Then%>class="Selected" style="display:block;"<%End If%>>
						<li id="now11" <%If divId=11 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=0&div=11" title="常规设置">常规设置</a></li>
						<li id="now12" <%If divId=12 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=placard&div=12" title="博客公告">博客公告</a></li>
						<li id="now13" <%If divId=13 Then%>class="Selected"<%End If%>><a href="user_friendurl.asp" title="博客友情连接">博客友情链接</a></li>
						<li id="now14" <%If divId=14 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=links&div=14" title="高级编辑友情链接">高级编辑友情链接</a></li>
						<li id="now15" <%If divId=15 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogpassword&div=15" title="加密博客">加密博客</a></li>
						<li id="now16" <%If divId=16 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogstar&div=16" title="申请博客之星">申请博客之星</a></li>
					</ul>
					<ul id="dleft_tab2" <%If divId=21 or divId=22 or divId=23 Then%>class="Selected" style="display:block;"<%End If%>>
						<li id="now21" <%If divId=21 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=userinfo&div=21" title="个人资料">个人资料</a></li>
						<li id="now22" <%If divId=22 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=userpassword&div=22" title="密码设置">密码设置</a></li>
						<li id="now23" <%If divId=23 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=userpassword&div=23" title="密码保护">密码保护</a></li>
					</ul>
					<ul id="dleft_tab3" <%If divId=31 or divId=32 or divId=33 Then%>class="Selected" style="display:block;"<%End If%>>
						<li id="now31" <%If divId=31 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogteam&div=31" title="团队成员管理">团队成员管理</a></li>
						<li id="now32" <%If divId=32 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogteam&div=32" title="我管理的团队">我加入的团队</a></li>
						<li id="now33" <%If divId=33 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogteam&div=33" title="我管理的团队">邀请朋友加入</a></li>
					</ul>
				</div>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
<%
	If oblog.ChkPost() = False Then
        oblog.adderrstr ("系统不允许从外部提交！")
    End If
Dim action
action = Request("action")
select Case action
    Case "blogteam"
        Call blogteam
    Case "links"
        Call links
    Case "savelinks"
        Call savelinks
    Case "placard"
        Call placard
    Case "saveplacard"
        Call saveplacard
    Case "blogpassword"
        Call blogpassword
    Case "addblogpassword"
        Call addblogpassword
    Case "unblogpassword"
        Call unblogpassword
    Case "savesitesetup"
        Call savesitesetup
    Case "blogstar"
        Call blogstar
    Case "saveblogstar"
        Call saveblogstar
    Case "saveuserlostpass"
        Call saveuserlostpass
    Case "userpassword"
        Call userpassword
    Case "saveuserpassword"
        Call saveuserpassword
    Case "saveuserinfo"
        Call saveuserinfo
    Case "userinfo"
        Call userinfo
	Case "aoboaccount"
		Call aoboaccount
    Case Else
        Call sitesetup
End select
%>

			</td>
		</tr>
	</tbody>
</table>
</body>
</html>
<%
Sub blogteam()
%>
				<div id="dTab31">
					<iframe class="FrmID" src="user_blogteam.asp?action=userteam&div=<%=divid%>" frameborder="0" scrolling="no"></iframe>
				</div>
<%
End Sub

Sub saveplacard()
    Dim rs, userplacard
    userplacard = oblog.filt_astr(Request.Form("edit"), 20000)
	If oblog.l_Group(15,0)=0 Then userplacard = FilterJS(userplacard)
	'-============================取消下面注释启用公告修改验证码验证
	'If Not oblog.codepass Then
	'		oblog.adderrstr ("验证码错误，请刷新后重新输入！")
	'		oblog.showusererr
	'		Response.end
	'end if
	If oblog.chk_badword(userplacard) >0 Then
		oblog.adderrstr ("站点公告中存在系统不允许的字符!")
		oblog.showusererr
		Exit Sub
	End if
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select user_placard from [oblog_user] where userid=" & oblog.l_uid, conn, 1, 3
    rs(0) = oblog.filtpath(userplacard)
    rs.Update
    rs.Close
    Dim blog
    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_placard oblog.l_uid
    Set rs = Nothing
    oblog.ShowMsg "修改公告成功", ""
End Sub

Sub placard()
Dim rs
Set rs = oblog.execute("select user_placard from [oblog_user] where userid=" & oblog.l_uid)
%>
				<div id="chk_idAll">
					<div id="dTab12">
						<form name="oblogform" method="post" action="user_setting.asp " <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
						<table class="Setting_Content" cellpadding="0">
							<tr>
								<td>
									你可以在这里放置你的照片，关于你的介绍，或者你愿意放上去的任何信息。
								</td>
							</tr>
							<tr>
								<td>
									<span id="loadedit"  style="display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> 正在载入编辑器...</span>
									<textarea id="edit" name="edit" style="display:none">
										<%=Server.HtmlEncode(OB_IIF(rs(0),""))%>
									</textarea>
<%If C_Editor_Type=2 Then  Server.Execute C_Editor & "/edit.asp" %>
								</td>
							</tr>
							<tr>
								<td><span style="display:none;float:left;width:470px;height:30px;">验证码:<input name="codestr" id="codestr" type="text"  size="4" maxlength="20" style="display:inline;height:18px;border:1px #1B76B7 solid;"><%=oblog.getcode%></span>
								<input name="Action" type="hidden" id="Action" value="saveplacard" />
								<input type="submit" name="Submit" id="Submit" value="提交修改" style="display:block;float:left;width:120px;height:50px;" />
								</td>
							</tr>
						</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "",1,"535","240"%>
<%
Set rs = Nothing
End Sub

Sub savelinks()
    Dim rs, links
    links = oblog.filt_astr(Request.Form("edit"), 20000)
	If oblog.l_Group(15,0)=0 Then links = FilterJS(links)
	If oblog.chk_badword(links) >0 Then
		oblog.adderrstr ("友情链接中存在系统不允许的字符!")
		oblog.showusererr
		Exit Sub
	End if
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select user_links from [oblog_user] where userid=" & oblog.l_uid, conn, 1, 3
    rs(0) = oblog.filtpath(links)
    rs.Update
    rs.Close
    Dim blog
    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_links oblog.l_uid
    Set rs = Nothing
    oblog.ShowMsg "修改友情连接成功", ""
End Sub

Sub links()
Dim rs
Set rs = oblog.execute("select user_links from [oblog_user] where userid=" & oblog.l_uid)
%>
				<div id="chk_idAll">
					<div id="dTab14">
						<form name="oblogform" method="post" action="user_setting.asp" <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
						<table class="Setting_Content" cellpadding="0">
							<tr>
								<td>
									你可以先输入文字或者图片，然后用 <img src="images/wlink.gif" align="absbottom" /> 按钮插入超级连接，推荐使用<a href="user_friendurl.asp">友情连接管理</a>。
								</td>
							</tr>
							<tr>
								<td>
									<span id="loadedit" style="display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> 正在载入编辑器...</span>
									<textarea id="edit" name="edit" style="display:none">
										<%=Server.HtmlEncode(OB_IIF(rs(0),""))%>
									</textarea >
<%If C_Editor_Type=2 Then  Server.Execute C_Editor & "/edit.asp" %>
								</td>
							</tr>
							<tr>
								<td>
								<input name="Action" type="hidden" id="Action" value="savelinks" />
								<input type="submit" name="Submit" id="Submit"  value="提交修改" />
								</td>
							</tr>
						</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "",1,"535","240"%>
<%
Set rs = Nothing
End Sub

Sub blogpassword()
%>
				<div id="chk_idAll">
					<div id="dTab15">
						<form name="oblogform" method="post" action="user_setting.asp?action=addblogpassword">
						<table class="Setting_Content" cellpadding="0">
<%
If oblog.l_Group(26,0)=1 Then
%>
							<tr>
								<td class="title">
									<label for="password">博客访问密码：</label>
								</td>
								<td>
									<form name="form1" method="post" action="user_setting.asp?action=" ><input type="password" id="password" name="password" /></br><input type="submit" name="Submit" value="全站加密" /></form>
									<span>加密后，你所有日志都需要通过密码验证后才能访问。</br>
	注意：设置完密码以后，需要<a href="user_update.asp" onclick="purl('user_setting.asp?action=userpassword&div=12','更新数据')">重新发布全站</a>！</span>
								</td>
							</tr>
<%Else%>
							<tr>
								<td class="title">
									系统禁止进行整站加密：
								</td>
								<td>
									如果之前您启用了整站加密功能，之前加密的内容仍能通过原来的方式访问。
								</td>
							</tr>
<%End If%>
							<tr>
								<td class="title">
								</td>
								<td>
									<form name="form2" method="post" action="user_setting.asp?action=unblogpassword" /><input type="submit" name="Submit" id="Submit" value="解除我站点的密码保护" /></form>
								</td>
							</tr>
						</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%end sub

Sub addblogpassword()
	If oblog.l_Group(26,0)=0 Then
		oblog.ShowMsg "系统不允许整站加密!",""
		Response.End
	End If
    Dim password, strtmp, blog
	password=Trim(Request("password"))
	if password="" then
		oblog.ShowMsg "密码不能为空!",""
		Response.End
	end if
    password = md5(password)
    oblog.execute("update [oblog_user] set blog_password='"&password&"' where userid="&oblog.l_uid)
    oblog.execute ("update [oblog_log] set blog_password=1 where userid=" & oblog.l_uid)
	oblog.execute ("update [oblog_album] set ishide=2 where userid=" & oblog.l_uid&" and ishide<>1")

    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_index 0
    blog.update_message 0
    Set blog = Nothing
	Session ("CheckUserLogined_"&oblog.l_uName)=""
	oblog.CheckUserLogined()
    oblog.ShowMsg "设置整站密码成功,请重新更新全站才可获得安全的加密保护！", ""
End Sub

Sub unblogpassword()
    Dim upath, blog
    oblog.execute ("update [oblog_user] set blog_password=null where userid=" & oblog.l_uid)
	oblog.execute ("update [oblog_album] set ishide=0 where userid=" & oblog.l_uid&" and ishide=2")
	oblog.execute ("update [oblog_log] set blog_password=0 where userid=" & oblog.l_uid)
	oblog.execute ("update [oblog_log] set isspecial=0 where userid=" & oblog.l_uid&" and ishide<>1 and (ispassword is null or ispassword='') and isneedlogin<>1 and (viewscores is null or viewscores=0) and (viewgroupid is null or viewgroupid=0 ) ")
    upath = Server.MapPath(oblog.l_udir)
    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_index 0
    blog.update_message 0
    Set blog = Nothing
	Session ("CheckUserLogined_"&oblog.l_uName)=""
	oblog.CheckUserLogined()
    oblog.ShowMsg "取消密码成功,请重新更新全站才可全部解密！", ""

End Sub

Sub sitesetup()
Dim rs, sstr, sublist, us, i,sstr1
If oblog.l_Group(6,0) = 0 Then sstr = "disabled"
If oblog.l_Group(7,0) = 0 Then sstr1 = "disabled"
Set rs = oblog.execute("select * from oblog_user where userid=" & oblog.l_uid)
us = rs("user_info")
If us = "" Or IsNull(us) Then
    sublist = 0
Else
    us = Split(us, "$")
    If us(0) <> "" Then sublist = CInt(us(0)) Else sublist = 0
End If

Dim user_domain,custom_domain
user_domain = oblog.filt_html(Trim(rs("user_domain")))
if true_domain=1 Then
	custom_domain = oblog.filt_html(Trim(rs("custom_domain")))
End if
If user_domain = "" Or IsNull(user_domain) Then
	sstr = ""
End If
If custom_domain = "" Or IsNull(custom_domain) Then
	sstr1 = ""
End If
%>
				<div id="chk_idAll">
					<div id="dTab11">
						<form name="oblogform" action="user_setting.asp" method="post">
						<table class="Setting_Content" cellpadding="0">
<%if Trim(oblog.cacheConfig(4))<>"" and oblog.cacheConfig(5) = "1" then%>
							<tr>
								<td class="title">
									<label for="user_domain">域名：</label>
								</td>
								<td>
									<input name="user_domain" id="user_domain" type="text" value="<%=user_domain%>" size="10" maxlength="20" <%=sstr%> /> <select name="user_domainroot" <%=sstr%>><%=oblog.type_domainroot(rs("user_domainroot"),0)%></select><input type="hidden" name="old_userdomain" value="<%=user_domain%>">
								</td>
							</tr>
<%end if%>
<%if true_domain=1 and oblog.l_Group(7,0) = "1" then%>
								<tr>
									<td class="title">
										<label for="custom_domain">绑定我的顶级域名：</label>
									</td>
									<td>
										<input name="custom_domain" id="custom_domain" type="text" value="<%=custom_domain%>" size="30" maxlength="50" <%=sstr1%> />
									<span>绑定前需确认域名ip已经解析到博客服务器。</span>
									</td>
								</tr>
<%end if%>
								<tr>
									<td class="title">
										<label for="blogname">站点名称：</label>
									</td>
									<td>
										<input name="blogname" id="blogname" type="text" value="<%=oblog.filt_html(rs("blogname"))%>" size="30" maxlength="20" />
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_classid">站点类别：</label>
									</td>
									<td>
										<select name="user_classid" id="user_classid" >
											<%=oblog.show_class("user",rs("user_classid"),0)%>
										</select>
									</td>
								</tr>
								<tr>
									<td class="title">
										允许将我加入博客团队：
									</td>
									<td>
										<label><input type="radio" value="1" name="en_blogteam" <%if rs("en_blogteam")<>0 then Response.write "checked"%> />是&nbsp;&nbsp;</label>
										<label><input type=radio value="0" name="en_blogteam" <%if rs("en_blogteam")=0 then Response.write "checked"%> />否</label>
										<span>允许别人将自己加入他创建的团队。</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										是否隐藏转向URL：
									</td>
									<td>
										<label><input type="radio" value="1" name="hideurl" <%if rs("hideurl")=1 then Response.write "checked"%> />是 &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="hideurl" <%if rs("hideurl")=0 Or  OB_IIF(rs("hideurl"),"")="" then Response.write "checked"%> />否</label>
										<span>开启后别人将看不到你的真实域名，只能看到你选择的二级域名。</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										分类日志是否以列表显示：
									</td>
									<td>
										<label><input type="radio" value="1" name="sublist" <%if sublist=1 then Response.write "checked"%> />是 &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="sublist" <%if sublist=0 then Response.write "checked"%> />否</label>
										<span>开启后打开你的文章分类导航将看到日志标题的排列，关闭则显示日志内容。</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										首页日志是否以列表显示：
									</td>
									<td>
										<label><input type="radio" value="1" name="indexlist" <%if rs("indexlist")=1 then Response.write "checked"%> />是 &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="indexlist" <%if rs("indexlist")=0 then Response.write "checked"%> />否</label>
										<span>开启后打开你的首页将看到最新日志标题的排列。</span>
										<span class="red">（要更新首页后才会生效）</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_showlogword_num">日志默认部分显示字数：</label>
									</td>
									<td>
										<input name="user_showlogword_num" id="user_showlogword_num" type="text"  value="<%=OB_IIF(rs("user_showlogword_num"),"500")%>" size="5" />
										<span>设置成0可显示全文。</span>
										<span class="red">（要更新首页后才会生效）</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_showlog_num">每页显示日志篇数：</label>
									</td>
									<td>
										<input name="user_showlog_num" id="user_showlog_num" type="text" id="user_showlog_num" value="<%=OB_IIF(rs("user_showlog_num"),"20")%>" size="5" />
										<span>首页显示日志数量，请不要设置为0或者太大的数字。</span>
										<span class="red">（要更新首页后才会生效）</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_photorow_num">每行显示相片数：</label>
									</td>
									<td>
										<input name="user_photorow_num" id="user_photorow_num" type="text" id="user_photorow_num" value="<%=OB_IIF(rs("user_photorow_num"),"4")%>" size="5" />
										<span>相册页面每行显示照片数。</span>
										<span class="red">（要更新首页后才会生效）</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_shownewcomment_num">显示最新回复条数：</label>
									</td>
									<td>
										<input name="user_shownewcomment_num" id="user_shownewcomment_num" type="text" value="<%=OB_IIF(rs("user_shownewcomment_num"),"8")%>" size="5" />
										<span class="red">（要更新首页后才会生效）</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_shownewlog_num">显示最新日志条数：</label>
									</td>
									<td>
										<input name="user_shownewlog_num" id="user_shownewlog_num" type="text" value="<%=OB_IIF(rs("user_shownewlog_num"),"8")%>" size="5" />
										<span class="red">（要更新首页后才会生效）</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_shownewmessage_num">显示最新留言条数：</label>
									</td>
									<td>
										<input name="user_shownewmessage_num" id="user_shownewmessage_num" type="text" value="<%=OB_IIF(rs("user_shownewmessage_num"),"8")%>" size="5" />
										<span class="red">（要更新首页后才会生效）</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										日志评论排列顺序：
									</td>
									<td>
										<label><input type="radio" value="1" name="comment_isasc" <%if rs("comment_isasc")=1 then Response.write "checked"%> />时间顺序 &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="comment_isasc" <%if rs("comment_isasc")=0 then Response.write "checked"%> />时间倒序</label>
									</td>
								</tr>
								<tr>
									<td class="title">
										编辑器类型：
									</td>
									<td>
										<label><input type="radio" value="2" name="isubbedit" <%if rs("isubbedit")=2 then Response.write "checked"%> />3.x版本(无法完美支持非IE浏览器) &nbsp;&nbsp;</label>
										<label><input type="radio" value="1" name="isubbedit" <%if rs("isubbedit")=1 then Response.write "checked"%> />4.x版本</label>
									</td>
								</tr>
								<tr>
									<td class="title">
										是否允许日志被其他用户推荐：
									</td>
									<td>
										<label><input type="radio" value="1" name="isdigg" <%if OB_iif(rs("isdigg"),1)=1 then Response.write "checked"%> />允许 &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="isdigg" <%if rs("isdigg")=0 then Response.write "checked"%> />不允许</label>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="siteinfo">站点简介：</label>
									</td>
									<td>
										<textarea name="siteinfo" id="siteinfo" cols="40" rows="5"><%=oblog.filt_html(rs("siteinfo"))%></textarea>
									</td>
								</tr>
								<tr>
									<td class="title">
									</td>
									<td>
										<input name="action" type="hidden" value="savesitesetup" />
										<input type="submit" id="Submit" value="保存修改" />
									</td>
								</tr>
							</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%
Set rs = Nothing
End Sub

Sub savesitesetup()
    Dim user_domain, user_domainroot, blogname, user_classid, en_blogteam, user_showlogword_num, user_showlog_num, user_shownewcomment_num, user_shownewlog_num, user_shownewmessage_num, hideurl, comment_isasc, siteinfo, user_photorow_num, custom_domain,isdigg
    Dim rs, blog
    user_domain = LCase(Trim(Request("user_domain")))
    user_domainroot = Trim(Request("user_domainroot"))
    blogname = EncodeJP(Trim(Request("blogname")))
    user_classid = Trim(Request("user_classid"))
    en_blogteam = Trim(Request("en_blogteam"))
    user_showlogword_num = Trim(Request("user_showlogword_num"))
    user_showlog_num = Trim(Request("user_showlog_num"))
    user_photorow_num = Trim(Request("user_photorow_num"))
    user_shownewcomment_num = Trim(Request("user_shownewcomment_num"))
    user_shownewlog_num = Trim(Request("user_shownewlog_num"))
    user_shownewmessage_num = Trim(Request("user_shownewmessage_num"))
    hideurl = Trim(Request("hideurl"))
    comment_isasc = Trim(Request("comment_isasc"))
    siteinfo = EncodeJP(Trim(Request("siteinfo")))
    custom_domain = Trim(Request("custom_domain"))
    isdigg = Trim(Request("isdigg"))
    If Trim(oblog.CacheConfig(4)) <> "" And oblog.CacheConfig(5) = 1 And oblog.l_Group(6,0) = 1 Then
        If user_domain = "" Or oblog.strLength(user_domain) > 20 Then oblog.adderrstr ("域名不能为空(不能大于14个字符)！")
        If user_domain <> Request("old_userdomain") And oblog.strLength(user_domain) < 4 Then oblog.adderrstr ("域名不能小于4个字符！")
        'If oblog.chk_regname(user_domain) Then oblog.adderrstr ("此域名系统不允许注册！")
        If oblog.chk_badword(user_domain) > 0 Then oblog.adderrstr ("域名中含有系统不允许的字符！")
        If oblog.chkdomain(user_domain) = False Then oblog.adderrstr ("域名不合规范，只能使用小写字母，数字！")
        If user_domainroot = "" Then oblog.adderrstr ("域名根不能为空！")
		If oblog.CheckDomainRoot(user_domainroot,0) = False Then oblog.adderrstr  ("域名根不合法！")
    End If
    If oblog.strLength(siteinfo) > 255 Then oblog.adderrstr ("站点简介不能大于255个字符！")
    If oblog.chk_badword(blogname) > 0 Then oblog.adderrstr ("blog名中含有系统不允许的字符！")
    If Not IsNumeric(user_showlogword_num) Then
        oblog.adderrstr ("日志默认部分显示字数必须为数字！")
    End If
	If oblog.CacheConfig(48)="1" Then
		Dim rsreg
		Set rsreg=oblog.execute("select Count(userid) From oblog_user Where blogname='" & ProtectSQL(blogname) & "' and userid<> " & oblog.l_uid)
    	If rsreg(0)>0 Then
    		oblog.adderrstr  ("您使用的博客名称: " & blogname & " 已被他人使用，请更换博客名称")
    	End If
    	rsreg.Close
	End If
    If Not IsNumeric(user_showlog_num) Then
        oblog.adderrstr ("每页显示日志数量必须为数字！")
    Else
        user_showlog_num = CLng(user_showlog_num)
        If user_showlog_num > 50 Then oblog.adderrstr ("每页显示日志数量必须小于50！")
    End If
    If Not IsNumeric(user_photorow_num) Then
        oblog.adderrstr ("每行显示相片数量必须为数字！")
    Else
        user_photorow_num = CLng(user_photorow_num)
        If user_photorow_num > 50 Then oblog.adderrstr ("每行显示相片数量必须小于50！")
    End If

    If Not IsNumeric(user_shownewcomment_num) Then
        oblog.adderrstr ("显示最新回复条数必须为数字！")
    Else
        user_shownewcomment_num = CLng(user_shownewcomment_num)
        If user_shownewcomment_num > 50 or user_shownewcomment_num<1 Then oblog.adderrstr ("显示最新回复条数不能大于50或者小于1！")
    End If

    If Not IsNumeric(user_shownewlog_num) Then
        oblog.adderrstr ("显示最新日志条数必须为数字！")
    Else
        user_shownewlog_num = CLng(user_shownewlog_num)
        If user_shownewlog_num > 50 or user_shownewlog_num<1 Then oblog.adderrstr ("显示最新日志条数不能大于50或者小于1！")
    End If

    If Not IsNumeric(user_shownewmessage_num) Then
        oblog.adderrstr ("显示最新留言条数必须为数字！")
    Else
        user_shownewmessage_num = CLng(user_shownewmessage_num)
        If user_shownewmessage_num > 50 or user_shownewmessage_num<1 then oblog.adderrstr ("显示最新留言条数不能大于50或者小于1！")
    End If
   If Trim(oblog.CacheConfig(4)) <> "" And oblog.CacheConfig(5) = 1 And oblog.l_Group(6,0) = 1 Then
        Set rs = oblog.execute("select userid from oblog_user where user_domain='" & oblog.filt_badstr(user_domain) & "' and user_domainroot='" & oblog.filt_badstr(user_domainroot) & "' and userid<>" & oblog.l_uid)
        If Not rs.EOF Or Not rs.bof Then oblog.adderrstr ("系统中已经有这个域名存在，请更改域名！")
    End If
    If true_domain = 1 And custom_domain <> "" Then
        If oblog.chk_badword(custom_domain) > 0 Then oblog.adderrstr ("绑定的顶级域名中含有系统不允许的字符！")
        Set rs = oblog.execute("select userid from oblog_user where custom_domain='" & oblog.filt_badstr(custom_domain) & "'" & " and userid<>" & oblog.l_uid)
        If Not rs.EOF Or Not rs.bof Then oblog.adderrstr ("系统中已经有其他人绑定了这个顶级域名，请更改域名或者联系管理员！")
    End If
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    If hideurl = "" Or IsNull(hideurl) Then hideurl = 0
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select * from oblog_user where userid=" & oblog.l_uid, conn, 1, 3
    If Not rs.EOF Then
		rs("blogname") = oblog.filt_astr(blogname, 50)
		If Trim(oblog.CacheConfig(4)) <> "" And oblog.CacheConfig(5) = 1 Then
			If oblog.l_Group(6,0) = 1 Or Trim (rs("user_domain")) = "" Or IsNull(rs("user_domain")) Then
				rs("user_domain") = user_domain
				rs("user_domainroot") = user_domainroot
			End if
        End If
		If true_domain = 1 Then
			If oblog.l_Group(7,0) = 1 Then
				rs("custom_domain") = custom_domain
            End If
        End If
        rs("user_classid") = user_classid
        rs("en_blogteam") = en_blogteam
        rs("user_showlogword_num") = user_showlogword_num
        rs("user_showlog_num") = user_showlog_num
        rs("user_photorow_num") = user_photorow_num
        rs("user_shownewcomment_num") = user_shownewcomment_num
        rs("user_shownewlog_num") = user_shownewlog_num
        rs("user_shownewmessage_num") = user_shownewmessage_num
        rs("hideurl") = hideurl
        rs("comment_isasc") = comment_isasc
        rs("siteinfo") = siteinfo
        rs("isubbedit") = OB_IIF(Request.Form("isubbedit"),2)
        rs("user_info") = CInt(Request.Form("sublist")) & "$0"
		rs("indexlist")=cint(Request.Form("indexlist"))
		rs("isdigg") = OB_IIF(isdigg,1)
        rs.Update
        rs.Close
    End If
    Set rs = Nothing
    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_blogname
    Set blog = Nothing
	Session ("CheckUserLogined_"&oblog.l_uName) = ""
	Oblog.CheckUserLogined
    oblog.ShowMsg "保存设置成功!", ""
End Sub

Sub saveblogstar()
    Dim rs, picurl, bloginfo, blogname
    picurl = Trim(Request("ico"))
    bloginfo = Trim(Request("bloginfo"))
    blogname = Trim(Request("blogname"))
    If picurl = "" Or oblog.strLength(picurl) > 250 Then oblog.adderrstr ("图片连接地址不能为空,且不能大于250个字符！")
    If bloginfo = "" Or oblog.strLength(bloginfo) > 250 Then oblog.adderrstr ("站点介绍不能为空,且不能大于250个字符！")
    If blogname = "" Or oblog.strLength(blogname) > 50 Then oblog.adderrstr ("博客名不能为空,且不能大于50个字符！")
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select top 1 * from oblog_blogstar Where userid=" & oblog.l_uid, conn, 1, 3
    If rs.EOF Then
        rs.addnew
        rs("userid") = oblog.l_uid
    End If
    rs("picurl") = picurl
    rs("info") = bloginfo
    If Trim(oblog.CacheConfig(4)) <> "" And oblog.CacheConfig(5) = 1  Then
		If oblog.l_Group(6,0) = 0 Then
			rs("userurl") = "http://" & oblog.l_udomain
		Else
			rs("userurl")=oblog.CacheConfig(3)&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext
		End if
    Else
        rs("userurl")=oblog.CacheConfig(3)&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext
    End If
    rs("blogname") = blogname
	rs("username") = Oblog.l_uName
	rs("usernickname") = Oblog.l_uNickname
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ShowMsg "提交完成，请等待管理员审核通过。", ""
End Sub

Sub blogstar()
    Dim rs, strTitle, strBlogName, strPicUrl, strBlogInfo, intState
    Set rs = Server.CreateObject("ADODB.RecordSet")
    rs.open "select * from oblog_blogstar Where userid=" & oblog.l_uid, conn, 1, 1
    If rs.EOF Then
        strTitle = "你目前还没有申请"
        intState = -1
    Else
        strPicUrl = rs("picurl")
        strBlogName = rs("blogname")
        strBlogInfo = rs("info")
        intState = rs("ispass")
        If intState = 1 Then
            strTitle = "你目前已经是博客之星，资料不可更改"
        Else
            strTitle = "你目前正在等待审核中，可以修改之前提交的资料"
        End If
    End If
    rs.Close
    Set rs = Nothing
%>
				<div id="chk_idAll">
					<div id="dTab21">
						<form name="oblogform" method="post" action="user_setting.asp?action=saveblogstar">
						<table  class="UserInfo" align="center" cellpadding="0" cellspacing="1">
							<tr>
								<td colspan="4">
									<div class="red"><%=strTitle%></div>
								</td>
							</tr>
							<tr>
								<td class="title">
									博客名字：
								</td>
								<td colspan="3">
									<input type="text" maxlength="50" size="60" name="blogname" value="<%=strBlogname%>" />
									<span>blog介绍请不要超过50字。</span>
								</td>
							</tr>
							<tr>
								<td class="title">用户头像：</td>
								<td colspan="3"><div class="user_face">
									<span><img src="<%=ProIco(strPicUrl,1)%>" class="face" id="imgIcon" width=<%=C_UserIcon_Width%> height=<%=C_UserIcon_Width%> /></span>
									<p><iframe id="d_file" frameborder="0" src="upload.asp?tMode=9&re=" width="400" height="30" scrolling="no"></iframe></p>
									<p>只支持jpg、gif、png，小于200k，默认尺寸为48px*48px<br /><br /></p>
									<p><select name="usertile" id="usertile" onchange="setusertile();"><option value="0">默认</option><%=GetUserTile%></select>　　<label>头像地址：<input name="ico"  id = "ico" type="text" value="<%=oblog.filt_html(strPicUrl)%>" size="60" maxlength="200"  onblur="getImg();" /></label></p>
								</div></td>
							</tr>
<!-- 							<tr>
								<td class="title">
									图片地址：
								</td>
								<td colspan="3">
									<input type="text" maxlength="250" size="60" name="picurl" value="<%=strPicUrl%>" />
									<span>图片地址可以放上你的照片，站点logo或者站点缩略图。<br />（图片尺寸最好缩小到130*100左右，以便管理员操作。）</span>
								</td>
							</tr>
 -->							<tr>
								<td class="title">
									blog介绍：
								</td>
								<td colspan="3">
									<textarea name="bloginfo" cols="50" rows="5"><%=strBlogInfo%></textarea><br />
									<span>填写Blog介绍和申请理由，审批通过后将公开显示，管理员有权对这些文字做适当调整。</span>
								</td>
							</tr>
							<tr>
								<td colspan="4" align="center">
<%
select Case intState
	Case -1
%>
									<input type="submit" id="Submit" value="提交申请资料" />
<%
Case 0
%>
									<input type="submit" id="Submit" value="修改申请资料" />
<%
Case 1
%>
									已经被确认为博客之星，资料不可更改，如果需要修改或撤销请与管理员联系。
<%End select%>
								</td>
							</tr>
						</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%
End Sub
Sub userpassword()
if is_ot_user=1 then
	Response.Redirect(ot_modifypass1)
end if
Dim rs
Set rs = oblog.execute("select question from oblog_user where userid=" & oblog.l_uid)
%>
				<div id="chk_idAll">
					<div id="dTab22"<%If divId=23 Then%>style="display:none;"<%End If%>>
						<form  method="post" action="user_setting.asp">
						<table class="Setting_Content" cellpadding="0">
							<tr>
								<td class="title">
									原始密码：
								</td>
								<td>
									<input name="oldpassword" type="password" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
									新密码：
								</td>
								<td>
									<input name="newpassword"  type="password" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
									重复密码：
								</td>
								<td>
									<input name="newpassword1" type="password" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
								</td>
								<td>
									<input name="action" type="hidden" value="saveuserpassword" />
									<input type="submit" id="Submit"  value=" 修改密码 " />
								</td>
							</tr>
						</table>
						</form>
					</div>
					<div id="dTab23" <%If divId=22 Then%>style="display:none;"<%End If%>>
						<form name="oblogform" action="user_setting.asp" method="post">
						<table class="Setting_Content" cellpadding="0">
							<tr>
								<td class="title">
									登录密码：
								</td>
								<td>
									<input name="password"  type="password" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
									密码提示问题：
								</td>
								<td>
									<input name="question"  type="text" size="30" maxlength="20" value="<%=oblog.filt_html(rs(0))%>">
								</td>
							</tr>
							<tr>
								<td class="title">
									找回密码答案：
								</td>
								<td>
									<input name="answer" type="text" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
								</td>
								<td>
									<input name="action" type="hidden" value="saveuserlostpass">
									<input type="submit" id="Submit" value="确认修改">
								</td>
							</tr>
						</table>
						</form>
					</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%
End Sub

Sub saveuserlostpass()
    Dim password, question, answer, rs
    password = Trim(Request("password"))
    question = Trim(Request("question"))
    answer = Trim(Request("answer"))
    If password = "" Then oblog.adderrstr ("错误：登录密码不能为空!")
    If question = "" Then oblog.adderrstr ("错误：提示问题不能为空！")
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select question,answer from oblog_user where userid="&oblog.l_uid&" and password='"&md5(password)&"'",conn,1,3
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("错误：登录密码输入错误!")
        oblog.showusererr
        Exit Sub
    Else
	    If API_Enable Then
	          Dim blogAPI
              Set blogAPI = New DPO_API_OBLOG
			  blogAPI.LoadXmlFile True
              blogAPI.UserName=oblog.l_uName
              blogAPI.PassWord=password
			  blogAPI.Question=Question
              blogAPI.Answer=Answer
		      Call blogAPI.ProcessMultiPing("update")
			  Set blogAPI=Nothing
		End If

        rs("question") = question
        If answer <> "" Then rs("answer") = md5(answer)
        rs.Update
        rs.Close
        Set rs = Nothing
        oblog.ShowMsg "修改找回密码资料成功！", ""
    End If

End Sub

Sub saveuserpassword()
    Dim oldpassword, newpassword, rs
    oldpassword = Trim(Request("oldpassword"))
    newpassword = Trim(Request("newpassword"))
    If oldpassword = "" Then oblog.adderrstr ("错误：原密码不能为空!")
    If newpassword = "" Or oblog.strLength(newpassword) > 14 Or oblog.strLength(newpassword) < 4 Then oblog.adderrstr ("错误：新密码不能为空(不能大于14小于4)！")
    If newpassword <> Trim(Request("newpassword1")) Then oblog.adderrstr ("错误：重复密码输入错误!")
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select password,TruePassWord from oblog_user where userid="&oblog.l_uid&" and password='"&md5(oldpassword)&"'",conn,1,3
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("错误：原密码输入错误!")
        oblog.showusererr
        Exit Sub
    Else

	    If API_Enable Then
			Dim blogAPI,j
			Set blogAPI = New DPO_API_OBLOG
			blogAPI.LoadXmlFile True
			blogAPI.UserName=oblog.l_uName
			blogAPI.PassWord=newpassword
			Call blogAPI.ProcessMultiPing("update")
			Set blogAPI=Nothing
			For j=0 To UBound(aUrls)
				strUrl=Lcase(aUrls(j))
				If Left(strUrl,7)="http://" Then
					Response.write("<script src="""&strUrl&"?syskey="&MD5(oblog.l_uName&oblog_Key)&"&username="&oblog.l_uName&"&password="&MD5(newpassword)&"&savecookie=0""></script>")
				End If
			Next
		End If

		Dim TruePassWord
		TruePassWord = RndPassword(16)
        rs("password") = md5(newpassword)
		rs("TruePassWord") = TruePassWord
        rs.Update
        rs.Close
        Set rs = Nothing
		oblog.savecookie oblog.l_uname, TruePassWord, 0
        oblog.ShowMsg "修改密码成功,下次需要重新登录！", ""
    End If
End Sub

Sub userinfo()
Dim rs
Set rs = oblog.execute("select * from oblog_user where userid=" & oblog.l_uid)
%>
				<div id="chk_idAll">
					<div id="dTab21">
						<form name="oblogform" action="user_setting.asp" method="post">
						<table  class="UserInfo" align="center" cellpadding="0" cellspacing="1">
							<tr>
								<td class="title">登录ＩＤ：</td>
								<td><span class="user_id"><%=rs("userName")%></span></td>
								<td class="title">用户等级：</td>
								<td class="userlevel"><span class="red"><%=oblog.l_Group(1,0)%></span></td>
							</tr>
							<tr>
								<td class="title">用户头像：</td>
								<td colspan="3"><div class="user_face">
									<span><img src="<%=ProIco(rs("user_icon1"),1)%>" class="face" id="imgIcon" width=<%=C_UserIcon_Width%> height=<%=C_UserIcon_Width%> /></span>
									<p><iframe id="d_file" frameborder="0" src="upload.asp?tMode=9&re=" width="400" height="30" scrolling="no"></iframe></p>
									<p>只支持jpg、gif、png，小于200k，默认尺寸为48px*48px<br /><br /></p>
									<p><select name="usertile" id="usertile" onchange="setusertile();"><option value="0">默认</option><%=GetUserTile%></select>　　<label>头像地址：<input name="ico"  id = "ico" type="text" value="<%=oblog.filt_html(rs("user_icon1"))%>" size="60" maxlength="200"  onblur="getImg();" /></label></p>
								</div></td>
							</tr>
							<tr>
								<td class="title"><label for="nickname">昵称：</label></td>
								<td colspan="3"><input name="nickname" id="nickname" type="text" value="<%=oblog.filt_html(rs("nickname"))%>" size="30" maxlength="20" /><input type="hidden" name="o_nickname" value="<%=oblog.filt_html(rs("nickname"))%>" /></td>
							</tr>
							<tr>
								<td class="title"><label for="truename">真实姓名：</label></td>
								<td><input name="truename" id="truename" type="text" value="<%=oblog.filt_html(rs("truename"))%>" size="30" maxlength="20" /></td>
								<td class="title">性别：</td>
								<td>
									<label><input type="radio" value="1" name="sex" <%if rs("Sex")=1 then Response.write "checked"%> />男</label>
									&nbsp;&nbsp;
									<label><input type="radio" value="0" name="sex" <%if rs("Sex")=0 then Response.write "checked"%> />女</label>
								</td>
							</tr>
							<tr>
								<td class="title"><label for="y">出生日期：</label></td>
								<td>
									<label><input value="<%=year(rs("birthday"))%>" name="y" id="y" size="2" maxlength="4" />年</label>
									<label><input value="<%=month(rs("birthday"))%>" name="m" size="2" maxlength="2" />月</label>
									<label><input value="<%=day(rs("birthday"))%>"  name="d" size="2" maxlength="2" />日</label>
								</td>
								<td class="title">省/市：</td>
								<td><%=oblog.type_city(rs("province"),rs("city"))%></td>
							</tr>
							<tr>
								<td class="title">职业：</td>
								<td><%oblog.type_job(rs("job"))%></td>
								<td class="title"><label for="Email">E-mail：</label></td>
								<td><input name="Email" id="Email" value="<%=oblog.filt_html(rs("userEmail"))%>" size="30" maxlength="50" /></td>
							</tr>
							<tr>
								<td class="title"><label for="homepage">主页：</label></td>
								<td colspan="3"><input maxlength="100" size="30" name="homepage" id="homepage" value="<%=oblog.filt_html(rs("Homepage"))%>" /></td>
							</tr>
							<tr>
								<td class="title"><label for="qq">QQ号码：</label></td>
								<td><input name="qq" id="qq" value="<%=oblog.filt_html(rs("qq"))%>" size="30" maxlength="20" /></td>
								<td class="title"><label for="msn">MSN：</label></td>
								<td><input name="msn" id="msn"value="<%=oblog.filt_html(rs("Msn"))%>" size="30" maxlength="50" /></td>
							</tr>
							<tr>
								<td class="title"><label for="tel">电话：</label></td>
								<td><input name="tel" id="tel" value="<%=oblog.filt_html(rs("tel"))%>" size="30" maxlength="50" /></td>
								<td class="title"><label for="address">通信地址：</label></td>
								<td><input name="address" id="address" value="<%=oblog.filt_html(rs("address"))%>" size="30" maxlength="250" /></td>
							</tr>
<%
If oblog.CacheConfig(51)="1" Then
If oblog.l_Group(34,0)="1" Then
%>
							<tr>
								<td colspan="4"><font class="red"><strong>本站允许通过邮件和手机发布日志</strong></font><br />在此处设置的邮箱或手机后,您可以通过这两种方式将需要发布的内容发送到<font class="red"><%=oblog.CacheConfig(52)%></font>,系统将自动解析内容并进行发布<br/>此处的邮箱地址与手机号码均不在站点公开显示，手机号码目前只支持中国移动GSM号码,即135~139号段</td>
							</tr>
							<tr>
								<td class="title"><label for="postmail">邮箱地址：</label></td>
								<td><input name="postmail" id="postmail" value="<%=oblog.filt_html(rs("postmail"))%>" size="30" maxlength="100" /></td>
								<td class="title"><label for="postmobile">手机号码：</label></td>
								<td><input name="postmobile" id="postmobile" value="<%=oblog.filt_html(rs("postmobile"))%>" size="30" maxlength="13" /></td>
							</tr>
<%
End If
End If
%>
							<tr>
								<td colspan="4" align="center">
									<input name="action" type="hidden" value="saveuserinfo" />
									<input type="submit" id="Submit" value="更新个人资料" />
								</td>
							</tr>
						</table>
						</form>
					</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%
Set rs = Nothing
End Sub

Sub saveuserinfo()
    Dim rs, nickname, email, birthday,usericon
    nickname = Trim(Request("nickname"))
    email = Trim(Request("email"))
    birthday = Trim(Request("y")) & "-" & Trim(Request("m")) & "-" & Trim(Request("d"))
    usericon=Trim(Request("ico"))
	If InStr(nickname,"$$$") > 0 Then oblog.adderrstr ("此昵称系统不允许注册！")
    'If oblog.chk_regname(nickname) Then oblog.adderrstr ("此昵称系统不允许注册！")
    If oblog.chk_badword(nickname) > 0 Then oblog.adderrstr ("昵称中含有系统不允许的字符！")
    If oblog.strLength(nickname) > 50 Then oblog.adderrstr ("昵称不能不能大于50字符！")
    '昵称唯一性判断
    If oblog.cacheConfig(47) = "1" And nickname <> "" And nickname <> Trim(Request("o_nickname")) Then
        Set rs = oblog.execute("select userid from oblog_user where nickname='" & ProtectSQL(nickname) & "'")
        If Not rs.EOF Or Not rs.bof Then oblog.adderrstr ("系统中已经有这个昵称存在，请更改昵称！")
    End If
    If birthday = "--" Then
        birthday = ""
    Else
        If Not IsDate(birthday) Then oblog.adderrstr ("生日日期格式错误！")
        If CLng(Trim(Request("y"))) > 2060 Then oblog.adderrstr ("生日年份过大！")
        If CLng(Trim(Request("y"))) < 1900 Then oblog.adderrstr ("生日年份过小！")
    End If
    If email = "" Then oblog.adderrstr ("电子邮件地址不能为空！")
    If Not oblog.IsValidEmail(email) Then oblog.adderrstr ("电子邮件地址格式错误！")
    If oblog.cacheConfig(22) = "1" Then
        Set rs = oblog.execute("select COUNT(userid) from oblog_user where useremail='" & ProtectSQL(email) & "' and userid<>" & oblog.l_uid)
        If rs(0) > 0 Then
			oblog.adderrstr ("系统中已经有这个Email存在，请更改Email！")
		End If
		rs.close
    End If
	email = Replace(email,"－","-")
    Dim rsreg
	Set rsreg=Nothing
	'---------------------------------------
		'Plus: Mail to Blog Start
		Dim sMail,sMobile,rstMail
		If oblog.cacheconfig(51)="1" Then
			If oblog.l_Group(34,0)="1" Then
				sMail=Trim(Request("postmail"))
				If  sMail<>"" Then
					if not oblog.IsValidEmail(sMail) then oblog.adderrstr("发布邮箱地址格式错误！")
				End If
				sMobile=Trim(Request("postmobile"))
				If  sMobile<>"" Then
					If Len(sMobile) = 11 And IsNumeric(sMobile) Then
						If CInt(Left(sMobile, 3)) >= 134 And CInt(Left(sMobile, 3)) <= 139 Or CInt(Left(sMobile, 3)) = 159 Then
							'bMobile = True
						Else
							oblog.adderrstr("您输入的手机号码错误或者系统暂不支持！")
						End If
					Else
						oblog.adderrstr("您输入的手机号码错误或者系统暂不支持！")
					End If
				End If

				set rstMail=Server.CreateObject("adodb.recordset")
				'判断Mail是否重复
				If  sMail<>"" Then
					rstMail.open "select * from oblog_user where postmail='" & LCase(Trim(sMail)) & "' And Userid<>" & oblog.l_uid,conn,1,1
					If Not rstMail.Eof Then
						oblog.adderrstr(sMail & " 已经被使用,请更换发布邮箱!" )
					End If
					rstMail.Close
				End If
				'判断手机号码是否重复
				If  sMobile<>"" Then
					rstMail.open "select * from oblog_user where postMobile='" & sMobile & "' And Userid<>" & oblog.l_uid,conn,1,1
					If Not rstMail.Eof Then
						oblog.adderrstr(sMobile & " 已经被使用,请更换发布号码!" )
					End If
					rstMail.Close
				End If
			End If
		End if
		Set rstMail=Nothing
		'Plus: Mail to Blog End
		'---------------------------------------
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select * from oblog_user where userid=" & oblog.l_uid, conn, 1, 3
    If Not rs.EOF Then
		If API_Enable Then
			  Dim blogAPI
			  Set blogAPI = New DPO_API_OBLOG
			  blogAPI.LoadXmlFile True
			  blogAPI.UserName=rs("UserName")
			  blogAPI.EMail=Email
			  blogAPI.Sex=CLng(Request("sex"))
			  blogAPI.QQ=oblog.filt_astr(Request("qq"),20)
			  blogAPI.MSN=oblog.filt_astr(Request("msn"),50)
			  blogAPI.truename=oblog.filt_astr(Request("truename"),20)
			  If birthday <> "" Then blogAPI.birthday=birthday
			  blogAPI.telephone=oblog.filt_astr(Request("tel"),50)
			  blogAPI.homepage=oblog.filt_astr(Request("homepage"),100)
			  blogAPI.province=oblog.filt_astr(Request("province"),18)
			  blogAPI.city=oblog.filt_astr(Request("city"),18)
			  blogAPI.address=oblog.filt_astr(Request("address"),250)
			  Call blogAPI.ProcessMultiPing("update")
			  Set blogAPI=Nothing
		End If
        rs("nickname") = oblog.filt_astr(nickname, 50)
        rs("truename") = oblog.filt_astr(Request("truename"), 20)
        rs("sex") = CLng(Request("sex"))
        rs("province") = oblog.filt_astr(Request("province"),20)
        rs("city") = oblog.filt_astr(Request("city"),20)
        If birthday <> "" Then rs("birthday") = oblog.filt_astr(birthday,20)
        rs("job") = oblog.filt_astr(Request("job"),20)
        rs("useremail") = oblog.filt_astr(email, 50)
        rs("homepage") = oblog.filt_astr(Request("homepage"), 100)
        rs("qq") = oblog.filt_astr(Request("qq"), 20)
        rs("msn") = oblog.filt_astr(Request("msn"), 50)
        rs("tel") = oblog.filt_astr(Request("tel"), 50)
        rs("address") = oblog.filt_astr(Request("address"), 250)
        rs("user_icon1") = RemoveHtml(oblog.filt_html(oblog.filt_astr(Request("ico"),200)))
		If oblog.cacheconfig(51)="1" Then
			If oblog.l_Group(34,0)="1" Then
				rs("postmail")=oblog.filt_astr(Request("postmail"),100)
				rs("postmobile")=oblog.filt_astr(Request("postmobile"),13)
			End If
		End if
        rs.Update
        rs.Close
    End If
    Set rs = Nothing
	Session ("CheckUserLogined_"&oblog.l_uName) = ""
	Oblog.CheckUserLogined
    oblog.ShowMsg "保存资料成功!", ""
End Sub

Function GetUserTile()
	Dim i
	Dim strTemp
	For i = 1 To 12
		strTemp = strTemp &"<option value=""usertile"&i&".gif"">头像"&i&"</option>"
	Next
	GetUserTile = strTemp
	strTemp = ""
End Function
%>
