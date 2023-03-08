<!--#include file="user_top.asp"-->
<%
dim action,iPage,sGuide
Dim tableName,itype,tableName1,tSQL
Dim teamID
Action=trim(request("Action"))
teamID=trim(request("teamid"))
If teamID<>"" Then teamID = CLng(teamID)
If teamID <> "" Then
	tableName = "[oblog_teamskin]"
	tableName1 = "[oblog_team]"
	tSQL = "teamID="&teamID
	itype = 1
Else
	tableName = "[oblog_userskin]"
	tableName1 = "[oblog_user]"
	tSQL = "userid="&oblog.l_uid
	itype = 0
End if
iPage=12
%>
<%
'设置每页模板显示显示数目
If  teamID<>"" Then
	G_P_FileName="user_template.asp?teamid="&teamID&"&page="
Else
	G_P_FileName="user_template.asp?page="
End if
G_P_PerMax=12
'模板显示顺序,1:倒序,最新的在最前面;2:顺序,最老的在最前面
Const P_USER_TEMPLATE_ORDERBY=1
%>
<%

select case Action
	case "saveconfig"
		call saveconfig()
	case "modiconfig"
		call modiconfig()
	case "savedefault"
		call savedefault()
	case "modiviceconfig"
		call modiviceconfig()
	case "saveviceconfig"
		call saveviceconfig()
	case "bakskin"
		call bakskin()
	case "good"
		call good()
	case Else
		call showconfig()
End select
%>
</table>
</body>
</html>
<script language=javascript>
function VerIfySubmit()
{
	If (document.oblogform.edit.value == "")
     {
        alert("请输入模板内容!");
	return false;
     }
	return true;
}
function setbak()
{
	document.bakform.bak.value='bak';
}
function setrestore()
{
	document.bakform.bak.value='restore';
}

function skin_help(action,t){
	if (t == 0)
	{
		if (action==0){
			var str="<div style='height:200px;overflow:auto;z-index:999999;'>主模版标签<hr />$show_log$ 重要，此标记显示日志主体部分，包括评论等信息。<br />$show_placard$ 此标记显示用户公告。 <br />$show_calendar$ 此标记显示日历。 <br />$show_newblog$ 此标记显示最新日志列表。 <br />$show_comment$ 此标记显示最新回复列表。<br />$show_subject$ 此标记显示专题分类。 <br />$show_subject_l$ 此标记横向显示专题分类。<br />$show_newblog$ 此标记显示最新日志列表。<br />$show_newmessage$ 此标记显示最新留言列表。<br />$show_info$ 此标记显示Blog名称，统计信息等。 <br />$show_login$ 此标记显示登录窗口 <br />$show_links$ 此标记显示链接信息。<br />$show_blogname$ 此标记显示用户blog名称，若名称为空则显示用户id。<br />$show_search$ 此标记显示搜索表单。<br />$show_xml$ 此标记显示rss连接标志。<br />$show_blogurl$ 此标记显示博客链接。<br />$show_myfriend$ 此标签显示我的好友。<br />$show_mygroups$ 此标签显示我加入的群组。<br />$show_photo$ 此标签调用相册。</div>";
		}else{
			var str="<div style='height:150px;overflow:auto;z-index:999999;'>副模版标签<hr /> $show_topic$ 此标记显示日志题目。 <br />$show_loginfo$ 此标记显示日志作者，发表时间等信息。 <br />$show_logtext$ 此标记显示日志正文。 <br />$show_more$ 此标记显示阅读全文，引用等链接。 <br />$show_emot$ 此标记仅显示显示表情图标。<br />$show_author$ 此标记仅显示作者名。<br />$show_addtime$ 此标记仅显示发表时间。<br />$show_topictxt$ 此标记仅显示日志标题。</div>";
		}
	}
	else
	{
		if (action==0){
			var str="<div style='height:200px;overflow:auto;z-index:999999;'><%=oblog.CacheConfig(69)%>主模版标签<hr />$group_id$ <%=oblog.CacheConfig(69)%>ID<br />$group_posts$ 最新文章<br /> $group_ico$  <%=oblog.CacheConfig(69)%>标记图片 <br /> $group_url$ <%=oblog.CacheConfig(69)%>访问地址 <br />$group_guide$ 导航链接<br /> $group_name$ <%=oblog.CacheConfig(69)%>名字 <br /> $group_creater$ <%=oblog.CacheConfig(69)%>创建人 <br /> $group_bottom$ 版权标识<br /> $group_comments$ 最近评论  <br />$group_placard$ 公告<br /> $group_links$ 友情链接 <br /> $group_info$ <%=oblog.CacheConfig(69)%>信息<br /> $group_bestuser$ 活跃用户<br /> $group_newuser$ 最新加入用户<br />  $group_admin$ 管理员信息<br /> $group_bestposts$ 精华帖子 <br />$group_photo$ <%=oblog.CacheConfig(69)%>相片 </div>";
		}else{
			var str="<div style='height:150px;overflow:auto;z-index:999999;'><%=oblog.CacheConfig(69)%>副模版标签<hr /> $group_list$ 内容标签 <br /> $group_post_title$ 帖子标题 <br />  $group_content$ 帖子内容<br />  $group_post_userico$ 作者头像<br />  $group_post_user$ 帖子作者 <br />  $group_post_time$ 发布时间 <br />  $group_post_content$ 帖子正文 <br />  $group_post_id$ 帖子ID <br /> $group_post_replys$ 回复按钮 <br />  $group_post_userurl$ 帖子作者地址 <br />   $group_post_high$ 帖子楼层  <br /> $group_post_m$帖子操作导航 <br /> </div>";
		}
	}
	 var oDialog = new dialog("<%=blogurl%>");
	 oDialog.init();
	 oDialog.event(str,'');
	 oDialog.button('dialogOk',"");
	 document.getElementById("ob_boxface").style.display="none";
}

</script>
<%
sub showconfig()
	Dim rs,SkinStrings,defaultskin,rst,sql,sClasses,sqlclass,classid
	Dim bookmark,sOrderby
	Set rs=Server.CreateObject("Adodb.Recordset")
	classid=Request("classid")
	If classid<>"" Then classid=Int(classid)
	If classid=0 Then classid=""
	'取默认模板
	rs.Open "select defaultskin from "&tableName1&" where "&tSQL,conn,1,3
	defaultskin=OB_IIF(rs(0),0)
	rs.Close
	If classid<>"" Then sqlclass=" And classid=" & classid:G_P_FileName="user_template.asp?teamid="&teamid&"&classid="&classid&"&page="
	'取用户/圈子模板可用分类
	If OB_IIF(teamid,"")="" Then
		If OB_IIF(oblog.l_Group(13,0),"")="" Then
			sql="select * from "&tableName&" where ispass=1 "  & sqlclass & " Order By Id"
			Set rst=oblog.Execute("Select * From oblog_skinclass Where itype="&itype&" And icount>0")
		Else
			sql="select * from "&tableName&" where ispass=1 And classid in (" & oblog.l_Group(13,0) &") "  & sqlclass & " Order By Id"
			Set rst=oblog.Execute("Select * From oblog_skinclass Where itype="&itype&" And classid in (" & oblog.l_Group(13,0) &")  And icount>0")
		End If
	Else
		If OB_IIF(oblog.l_Group(5,0),"")="" Then
			sql="select * from "&tableName&" where ispass=1 "  & sqlclass & " Order By Id"
			Set rst=oblog.Execute("Select * From oblog_skinclass Where itype="&itype&" And icount>0")
		Else
			sql="select * from "&tableName&" where ispass=1 And classid in (" & oblog.l_Group(5,0) &") "  & sqlclass & " Order By Id"
			Set rst=oblog.Execute("Select * From oblog_skinclass Where itype="&itype&" And classid in (" & oblog.l_Group(5,0) &")  And icount>0")
		End If
	End If
	sClasses="<option value=""0"">查看全部模板</option>"
	Do While Not rst.Eof
		sClasses=sClasses & "<option value=" & rst("classid") & ">" & rst("classname")& "(有" & rst("icount") &"个模板)</option>"
		rst.Movenext
	Loop
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
<!--					<fieldset id="Template" class="FieldsetForm">
						<legend>当前模板：</legend>
							<ul>
								<li>
								<%
								'If defaultskin=0 Then
									'Response.Write "									<strong class=""red"">你当前还没有选择任何模板</strong>" & VBCRLF
								'Else
									'Set rst=oblog.Execute("Select * From "&tableName&" Where id=" & defaultskin)
									'If Not rst.Eof Then
										'Response.Write "<ul class=""Skin_onmouseover"">" & VBCRLF
										'Response.Write "										<li class=""l1""><a href=""showskin.asp?id=" & rst("id") & """ target=""_blank""><img src="""&oblog.filt_html(rst("skinpic"))&""" title=""点击预览"" width=""200"" height=""122"" border=""0"" /></a></li>" & VBCRLF
										'Response.Write "										<li class=""l2"">名称：" & "<strong>" & rst("userskinname") &"</strong></li>" & VBCRLF
										'Response.Write "										<li class=""l3"">作者：<a href=""" & rst("skinauthorurl") &""" target=""_blank"">" & "<span class=""blue"">" & rst("skinauthor") &"</span></a></li>" & VBCRLF
										'Response.Write "									</ul>" & VBCRLF
									'End If
									'Set rst=Nothing
								'End If
								%>
								</li>
							</ul>
					</fieldset>-->
<%
Dim lPage,lAll,lPages
If P_USER_TEMPLATE_ORDERBY=1 Then sOrderby=" Desc"
rs.Open sql & sOrderby,conn,1,3
lAll=clng(rs.recordcount)
    If lAll=0 Then
    	rs.Close
    	Set rs=Nothing
    	%>
				<div id="chk_idAll">
					<!-- 没有相关记录 -->
					<div class="msg"><%=sGuide & " 没有相关纪录" %></div>
					<!-- 没有相关记录 end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
    	<%
    	Exit Sub
    End If

	'分页
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = clng(Request("page"))
	End If

	'设置缓存大小 = 每页需显示的记录数目
	rs.CacheSize = iPage
	rs.PageSize = iPage
	rs.movefirst
	lPages = rs.PageCount
	If lPage>lPages Then lPage=lPages
	rs.AbsolutePage = lPage
%>

					<fieldset id="Template" class="FieldsetForm">
						<legend>选择模板：<span class="red">（<%
								If defaultskin=0 Then
									Response.Write "<strong class=""red"">你当前还没有选择任何模板</strong>" & VBCRLF
								Else
									Set rst=oblog.Execute("Select * From "&tableName&" Where id=" & defaultskin)
									If Not rst.Eof Then
										If teamid <>"" Then
											tSQL = "teamskinid=" & rst("id")
										Else
											tSQL = "id=" & rst("id")
										End if
										Response.Write "当前模板为：<a href=""showskin.asp?" & tSQL & """ target=""_blank""><strong>" & rst("userskinname") &"</a></strong>" & VBCRLF
									End If
									Set rst=Nothing
								End If
								%>）</span></legend>
							<ul>
								<li id="form">
								<%If teamid = "" Then %>
									<form id="sClasses" name="formclass" method="post" action="user_template.asp">
										<select name="classid" onchange="javascript:window.location='user_template.asp?action=showconfig&classid='+this.options[this.selectedIndex].value;">
												<option>请选择模板分类</option>
												<%=sClasses%>
										</select>
									</form>
								<%End if%>
									<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
								</li>
								<li>
									<%
										SkinStrings = SkinStrings & GetSkinList(rs,defaultskin)
										rs.Close
										set rs=Nothing
										Response.Write SkinStrings
										SkinStrings=""
									%>
									<input type="hidden" value="<%=request("u")%>">
									<input type="hidden" name="teamid" value="<%=teamid%>">
								</li>
							</ul>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%
	set rs=nothing
End sub

sub savedefault()
	dim rs,rsskin,isdefaultID
	isdefaultID=clng(trim(request("radiobutton")))
	set rsskin=oblog.execute("select skinmain,skinshowlog from "&tableName&" where id="&isdefaultID)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select user_skin_main,user_skin_showlog,defaultskin from "&tableName1&" where "&tSQL,conn,1,3
	rs(0)=rsskin(0)
	rs(1)=rsskin(1)
	rs(2)=isdefaultID
	set rsskin=nothing
	rs.update
	rs.close
	set rs=nothing
	'如果是新手模式
	If teamid = "" Then
		If oblog.l_uNewbie=1 Then
			'创建新的用户目录
			If oblog.CacheConfig(59) = "0" Then oblog.CreateUserDir oblog.l_uname, 1
			oblog.execute "Update oblog_user Set newbie=0 Where userid=" & oblog.l_uid
			'清空Session，防止选择后无法即时生效
			Session ("CheckUserLogined_"&oblog.l_uName) = ""
			Oblog.CheckUserLogined()
			updateindex()
			oblog.ShowMsg "模板选择成功！","user_template.asp?action=showconfig&teamid="&teamid
		Else
'			Session ("CheckUserLogined_"&oblog.l_uName) = ""
			updateindex()
			oblog.ShowMsg "修改成功,首页已经更新，其他页面请手动更新！","user_template.asp?action=showconfig&teamid="&teamid
		End If
	Else
		oblog.ShowMsg "修改成功！","user_template.asp?action=showconfig&teamid="&teamid
	End if
End sub

sub saveconfig()
	dim rs,sql,sContent,iChk
	If oblog.l_Group(14,0)=0 Then
		oblog.ShowMsg "你所在的组不允许修改模板！","user_template.asp"
		Exit Sub
	End If
	set rs=server.CreateObject("adodb.recordset")
	sql="select user_skin_main from "&tableName1&" where "&tSQL
	rs.open sql,conn,1,3
	sContent=oblog.filtpath(oblog.filt_badword(Trim(request("edit"))))
'	OB_DEBUG sContent,1
	'内容检查
	iChk=oblog.chk_badword(sContent)
	If iChk>0 Then
		oblog.ShowMsg "模板中存在系统禁止的字符!",""
		Response.End
	End If
	sContent=Replace(Replace(sContent,"<%","&lt;%"),"%"&">","%&gt;")
	rs(0)=sContent
	'脚本过滤
	If oblog.l_Group(15,0)=0 Then rs(0)=oblog.CheckScript(sContent)

	rs.update
	rs.close
	Set rs=Nothing
	sContent=""
	If teamID <> "" Then
		oblog.ShowMsg "修改成功！",""
	Else
		updateindex()
		oblog.ShowMsg "修改成功,首页已经更新，其他页面请手动更新！",""
	End If
End sub

sub saveviceconfig()
	dim rs,sql,sContent
	Dim iChk
	If oblog.l_Group(14,0)=0 Then
		oblog.ShowMsg "你所在的组不允许修改模板！",""
		Exit Sub
	End If
	set rs=server.CreateObject("adodb.recordset")
	sql="select user_skin_showlog from "&tableName1&" where "&tSQL
	rs.open sql,conn,1,3
	sContent=oblog.filtpath(oblog.filt_badword(trim(request("edit"))))
	'内容检查
	iChk=oblog.chk_badword(sContent)
	If iChk>0 Then
		oblog.ShowMsg "模板中存在系统禁止的字符!",""
		Response.End
	End If
	'脚本过滤
	If oblog.l_Group(15,0) = "0" Then rs(0)=oblog.CheckScript(sContent)
	rs(0)=sContent
	rs.update
	rs.close
	set rs=nothing
	sContent=""
	If teamID <> "" Then
		oblog.ShowMsg "修改成功！",""
	Else
		updateindex()
		oblog.ShowMsg "修改成功,首页已经更新，其他页面请手动更新！",""
	End If
End sub

sub modiconfig()
	If oblog.l_Group(14,0)=0 Then
		oblog.ShowMsg "你所在的组不允许修改模板！","user_template.asp"
		Exit Sub
	End If
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_main,defaultskin from "&tableName1&" where "&tSQL)
	If rs("defaultskin")=0  or rs("defaultskin")="" Or IsNull (rs(1)) Then
		set rs=nothing
		oblog.adderrstr("请先选择一个默认模板！")
		oblog.showusererr
		exit sub
	End If
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUp" class="FieldsetForm">
						<legend>修改主模板：</legend>
						<form method="POST" action="user_template.asp" id="oblogform" name="oblogform"  <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
							<ul>
								<li><strong>主模板决定页面整体风格</strong>，建议修改前先<a href="user_template.asp?action=bakskin&teamid=<%=teamid%>"><span class="blue">备份模板</span></a>。<br /><a href="#" onclick="skin_help(0,<%=OB_IIF(teamid,0)%>);" ><span class="blue">主模板标签说明</span></a></li>
								<li><span id="loadedit" class="red" style="display:<%=C_Editor_LoadIcon%>;"><img src="images/loading.gif" align="absbottom" /> 正在载入编辑器...</span>
									<textarea id="edit" name="edit" style="	display:none"><%=Server.HtmlEncode(rs(0))%></textarea>
										<%If C_Editor_Type=2 Then  Server.Execute C_Editor & "/edit.asp" %>
								</li>
								<input name="teamid" type="hidden" id="teamid" value="<%=teamid%>" />
								<li><input name="Action" type="hidden" id="Action" value="saveconfig" /><input name="cmdSave" type="submit" id="Submit" value=" 保存修改 " style="height:30px;" /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "edit",1,"580","300"%>
<%
set rs=nothing
End sub

sub modiviceconfig()
	If oblog.l_Group(14,0)=0 Then
		oblog.ShowMsg "你所在的组不允许修改模板！","user_template.asp"
		Exit Sub
	End If
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_showlog,defaultskin from "&tableName1&" where "&tSQL)
	If rs("defaultskin")=0  or rs("defaultskin")="" Or IsNull(rs(1)) Then
		set rs=nothing
		set rs=nothing
		oblog.adderrstr("请先选择一个默认模板！")
		oblog.showusererr
		exit sub
	End If
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll" style="overflow-y:auto;">
					<fieldset id="BackUp" class="FieldsetForm">
						<legend>修改副模板：</legend>
						<form method="POST" action="user_template.asp?id=<%=clng(request.QueryString("id"))%>" id="oblogform" name="oblogform" <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
							<ul>
								<li><strong>副模板决定日志部分显示风格</strong></a>。<br /><a href="#" onclick="skin_help(1,<%=OB_IIF(teamid,0)%>);"><span class="blue">副模板标签说明</span></a></li>
								<li><span id="loadedit" class="red" style="display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> 正在载入编辑器...</span>
									<textarea id="edit" name="edit" style="display:none;"><%=Server.HtmlEncode(rs(0))%></textarea>
									<%If C_Editor_Type=2 Then  Server.Execute C_Editor & "/edit.asp" %>
									</li>
								<li><input name="action" type="hidden" id="action" value="saveviceconfig" />
								<input name="teamid" type="hidden" id="teamid" value="<%=teamid%>" /><input type="submit"  value="保存修改" id="Submit" /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "edit",1,"580","300"%>
<%
	set rs=nothing
End sub

sub bakskin()
	dim bak,rs
	bak=request("bak")
	If bak="bak" Then
		oblog.execute("update "&tableName1&" set bak_skin1=user_skin_main,bak_skin2=user_skin_showlog where "&tSQL)
		oblog.ShowMsg "备份模板成功！",""
	ElseIf bak="restore" Then
		set rs=oblog.execute("select bak_skin1,bak_skin2 from "&tableName1&" where "&tSQL)
		If rs(0)="" or rs(1)="" or isnull(rs(0)) or isnull(rs(1)) Then
			set rs=nothing
			oblog.adderrstr("所备份的模板为空，不允许恢复！")
			oblog.showusererr
		End If
		oblog.execute("update "&tableName1&" set user_skin_main=bak_skin1,user_skin_showlog=bak_skin2 where "&tSQL)
		set rs=nothing
		oblog.ShowMsg "恢复模板成功！",""
	End If
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUpTemplate" class="FieldsetForm">
						<legend>备份模板</legend>
						<form name="bakform" method="post" action="user_template.asp?action=bakskin&teamid=<%=teamid%>">
							<ul>
								<li><input type="hidden" name="bak" id="bak" value="" /><input type="submit" id="Submit" name="Submit" value="备份模板" onclick="setbak();" /></li>
								<li>备份现在使用中的模板。<a href="showskin.asp?<%=tSQL%>" target ="_blank"><span class="red">查看目前模板备份</span></a></li>
								<li></li>
							</ul>
					</fieldset>
					<fieldset id="BackUpTemplate" class="FieldsetForm">
						<legend>恢复模板</legend>
							<ul>
								<li><input type="submit" id="Submit" name="Submit2" value="恢复模板" onclick="setrestore();" /></li>
								<li>恢复成您最后备份的模板，将覆盖现在使用中的模板。</li>
								<li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%End Sub

sub good()
	dim good,rs,rsu,skinname
	good=request("good")
	If good="save" Then
		skinname=request("skinname")
		If skinname="" or oblog.strLength(skinname)>50  Then oblog.adderrstr("用户名不能为空(不能大于50)！")
		If oblog.errstr<>"" Then oblog.showusererr:exit sub
		set rsu=oblog.execute("select user_skin_main,user_skin_showlog from "&tableName1&" where "&tSQL)
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select top 1 * from ["&tableName&"]",conn,1,3
		rs.addnew
		rs("userskinname")=skinname
		rs("skinmain")=rsu(0)
		rs("skinshowlog")=rsu(1)
		rs("skinauthor")=oblog.l_uname
		rs("skinauthorurl")=oblog.l_udir&"/"&oblog.l_uid&"/index."&f_ext
		rs.update
		rs.close
		set rsu=nothing
		set rs=nothing
		oblog.ShowMsg "推荐成功，请等待管理员审核！",""
	End If
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUpTemplate" class="FieldsetForm">
						<legend>推荐我的模板</legend>
						<form name="good" action="user_template.asp?action=good&good=save" method="post">
							<ul>
								<li>如果你的模板很漂亮，推荐给管理员，可以放在模板数据库里，让更多人使用哦!<br />模板会注明作者和你的blog连接。请不要提交已经存在的用户模板！</li>
								<li>模板名称：<input name="skinname" type="text" value="" size="20" maxlength="20" /></li>
								<li><input type="submit" id="Submit" value="推荐" /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%End sub
sub updateindex()
	dim blog
	set blog=new class_blog
	blog.userid=oblog.l_uid
	blog.update_index 0
	blog.update_message 0
	blog.CreateFunctionPage
	set blog=nothing
End sub

Function GetSkinList(byref rst,byval defaultskin)
	Dim strSkins,ustr,iCount
	iCount=0
	strSkins=""
	Do While Not rst.eof
		If teamid <>"" Then
			tSQL = "teamskinid=" & rst("id")
		Else
			tSQL = "id=" & rst("id")
		End If
		If rst("skinauthorurl")<>"" Then
			ustr="<a href="""&rst("skinauthorurl")&""" target=""_blank""><span class=""blue"">"&rst("skinauthor")&"</span></a>"
		Else
			ustr=rst("skinauthor")
		End If
		strSkins = strSkins & "									<ul onmouseover=""this.className='Skin_onmouseover'"" onmouseout=""this.className='Skin_onmouseout'"" class=""Skin_onmouseout"">" & VBCRLF
		strSkins = strSkins & "										<li class=""l1""><a href='showskin.asp?"&tSQL&" ' target=_blank>" & VBCRLF
		If rst("skinpic")="" or isnull(rst("skinpic")) Then
			strSkins = strSkins & "<img src=""images/nopic.gIf"" title=""对不起,该模板没有预览图"" width=""200"" height=""122"" border=""0"" />" & VBCRLF
		Else
			strSkins = strSkins & "<img src="""&oblog.filt_html(rst("skinpic"))&""" title=""点击预览"" width=""200"" height=""122"" border=""0"" />" & VBCRLF
		End If
		strSkins = strSkins & "</a></li>" & VBCRLF
		strSkins = strSkins & "										<li class=""l2"" title=""模板：" & rst("userskinname") & """>模板：<strong>" & rst("userskinname") & "</strong></li>" & VBCRLF
		strSkins = strSkins & "										<li class=""l3"" title=""作者："&rst("skinauthor")&""">作者：" & ustr & "</li>" & VBCRLF
'		strSkins = strSkins & "<div class=""skin_used""><a href=""user_template.asp?action=savedefault&teamid="&teamid&"&radiobutton=" & rst("id") & """ >应用此模板</a></div>"
		strSkins = strSkins & "										<li class=""l4""><input type=""submit"" value=""应用此模板"" onClick=""window.location='user_template.asp?action=savedefault&teamid="&teamid&"&radiobutton=" & rst("id") & "'"" /></li>"
		strSkins = strSkins & "									</ul>	" & VBCRLF & VBCRLF

		iCount = iCount+1
		if iCount >=iPage then exit do
		rst.movenext
	Loop
	GetSkinList = strSkins
End Function
%>
