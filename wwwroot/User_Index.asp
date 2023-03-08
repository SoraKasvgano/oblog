<!--#include file="user_top.asp"-->
<%
'是否启用短信息提示音
Const En_SMS_Sound=0
If API_Enable Then
	If session("turl")<>"" Then
		Dim arrturl,i,turl
		turl=Replace(session("turl"),"$","&")
		arrturl=Split(turl,"@@@")
		For i=0 To UBound(arrturl)
			Response.Write "<script language=JAVASCRIPT src="""&arrturl(i)&"""></script>" & vbcrlf
		Next
		Response.Flush
		session("turl")=""
	End if
End if
Dim MainUrl,UserUrl,rs,jscmd,tarr
MainUrl=Replace(Request.QueryString("url"),"$","&")
if instr(MainUrl,"user_post.asp") Then
	jscmd="go_cmdurl('发布日志','tab3')"
elseif instr(MainUrl,"user_url.asp") Then
	If InStr(MainUrl,"stitle=") > 0 Then
		tarr = Split (MainUrl,"stitle=")
		tarr(1) = Server.UrlEncode(tarr(1))
		MainUrl = tarr(0) & "stitle=" & tarr(1)
	End if
	jscmd="go_cmdurl('添加订阅','tab3')"
elseif instr(MainUrl,"user_friends.asp") then
	jscmd="go_cmdurl('我的好友','tab3')"
elseif instr(MainUrl,"User_myactions.asp") then
	jscmd="go_cmdurl('我的好友','tab3')"
elseif MainUrl="" then
	MainUrl="about:blank"
end if
If oblog.l_uNewbie=1 Then
	MainUrl="user_template.asp?action=showconfig"
	jscmd="go_cmdurl('选择模版','tab3')"
end if
'If Instr(MainUrl,"user_index.asp")>0 Then 	MainUrl="user_index_frame1.asp"
'取个人域名
If oblog.CacheConfig(5)=1 Then
	If Left(oblog.l_udomain,8)="http://." Or Trim(oblog.l_udomain)="." Then
		UserUrl="<a href="""&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext&""" target=""_blank"">我的首页</a>"
	Else
		UserUrl="<a href=""http://"&oblog.l_udomain&""" target=""_blank"">"&oblog.l_udomain&"</a>"
	End If
Else
	UserUrl="<a href="""&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext&""" target=""_blank"">我的首页</a>"
End If
If true_domain=1 and oblog.l_ucustomdomain<>"" then
	UserUrl="<a href=""http://"&oblog.l_ucustomdomain&""" target=""_blank"">"&oblog.l_ucustomdomain&"</a>"
End If

%>

<link rel="stylesheet" href="oBlogStyle/UserAdmin/7/default.css" type="text/css" />

<table id="IndexTableBody" cellpadding="0">
	<thead>
		<tr>
			<th>
				<div id="logo" title="用户管理后台">用户管理后台</div>
			</th>
			<th>
				<%=UserUrl%>&nbsp;|&nbsp;
				<a href="index.asp" target="_blank">站点首页</a>&nbsp;|&nbsp;
				<a href="user_pmmanage.asp" onClick="go_cmdurl('短消息',this)" target="content3">短消息<span id="ob_pm"></span></a>&nbsp;|&nbsp;
				<a href="user_setting.asp" onClick="go_cmdurl('博客设置',this)" target="content3">设置</a>&nbsp;|&nbsp;
				<a href="user_help.asp"  onclick="go_cmdurl('用户帮助',this)" target="content3">帮助</a>&nbsp;|&nbsp;
				<a href="user_index.asp?t=logout" class="txt_nor">退出</a>&nbsp;
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td class="menu">
				<ul class="bigbtu">
					<li id="now01"><a href="user_post.asp" onClick="go_cmdurl('发布日志',this);border_left('TabPage2','left_tab2');" title="发表日志" target="content3">发布日志</a></li>
					<li id="now02"><a href="user_post.asp?t=1&tMode=normal" onClick="go_cmdurl('发布相片',this);border_left('TabPage2','left_tab3');" title="发布相片" target="content3">发布相片</a></li>
				</ul>
			</td>
			<td class="tab">
			   <ul id="TabPage1">
					<li id="Tab1" class="Selected" onClick="javascript:switchTab('TabPage1','Tab1');" title="管理首页"><span>管理首页</span></li>
					 <%If oblog.CacheConfig(81) = "1" Then%>
					 <li id="Tab2" onClick="pmusicurl();" title="AOBO音乐盒" style="display:block" ><span style="display:block">音乐盒</span></li><%else%>
					 <li id="Tab2" onClick="javascript:switchTab('TabPage1','Tab2');" title="AOBO音乐盒"  ><span  style="display:block">音乐盒</span><li>

					 <%End If %>
					<li id="Tab3" onClick="javascript:switchTab('TabPage1','Tab3');"><span id="dnow99" style="display:block">空白页面</span></li>
			   </ul>
			</td>
		</tr>
		<tr>
			<td class="t1">
				<div id="contents">
					<table cellpadding="0">
						<tr class="t1">
							<td><div class="menu_top"></div></td>
						</tr>
						<tr class="t2">
							<td>
								<div class="menu" id="TabPage3">
									<ul id="TabPage2">
										<li id="left_tab1" class="Selected" onClick="javascript:border_left('TabPage2','left_tab1');" title="常用"><span>常用</span></li>
										<li id="left_tab2" onClick="javascript:border_left('TabPage2','left_tab2');" title="日志"><span>日志</span></li>
										<li id="left_tab3" onClick="javascript:border_left('TabPage2','left_tab3');" title="相册"><span>相册</span></li>
										<li id="left_tab4" onClick="javascript:border_left('TabPage2','left_tab4');" title="<%=oblog.CacheConfig(69)%>"><span><%=oblog.CacheConfig(69)%></span></li>
										<li id="left_tab7" onClick="javascript:border_left('TabPage2','left_tab7');getfeedlist();" title="订阅"><span>订阅</span></li>
										<li id="left_tab5" onClick="javascript:border_left('TabPage2','left_tab5');" title="文件"><span>文件</span></li>
										<li id="left_tab6" onClick="javascript:border_left('TabPage2','left_tab6');" title="模板"><span>模板</span></li>
										<%If oblog.CacheConfig(12) = "1" Then %>
										<li id="left_tab8" onClick="javascript:border_left('TabPage2','left_tab8');" title="服务"><span>服务</span></li>
										<%End if%>
									</ul>
									<div id="left_menu_cnt">
										<ul id="dleft_tab1">
											<li id="now11" class="Selected"><a href="user_blogmanage.asp" onClick="go_cmdurl('日志管理',this);" target="content3" title="日志管理"><span>日志管理</span></a></li>
											<li id="now12"><a href="user_post.asp?action=showphoto&t=1" onClick="go_cmdurl('浏览相册',this);" target="content3" title="浏览相册"><span>浏览相册</span></a></li>
											<li id="now13"><a href="user_comments.asp" onClick="go_cmdurl('日志评论',this);" target="content3" title="日志评论"><span>日志评论</span></a></li>
											<li id="now14"><a href="user_Albumcomments.asp" onClick="go_cmdurl('相册评论',this);" target="content3" title="相册评论"><span>相册评论</span></a></li>
											<li id="now15"><a href="user_messages.asp" onClick="go_cmdurl('访客留言',this)" target="content3" title="访客留言"><span>访客留言</span></a></li>
											<li id="now16"><a href="user_diggs.asp" onClick="go_cmdurl('推荐日志',this)" target="content3" title="推荐日志"><span>推荐日志</span></a></li>
											<li id="now17"><a href="user_friendurl.asp" onClick="go_cmdurl('友情链接',this);" target="content3" title="友情链接"><span>友情链接</span></a></li>
											<li id="now18"><a href="user_setting.asp?action=placard&div=12" onClick="go_cmdurl('博客公告',this);" target="content3" title="博客公告"><span>博客公告</span></a></li>
											<%if oblog.CacheConfig(17)=1 then%>
											<li id="now19"><a href="user_codes.asp" onClick="go_cmdurl('邀请码',this)" target="content3" title="可用邀请码"><span>可用邀请码</span></a></li>
											<%end if%>
											<li id="now1a"><a href="user_friends.asp" onClick="go_cmdurl('我的好友',this)" target="content3" title="我的好友"><span>我的好友</span></a></li>
											<li id="now1b"><a href="user_update.asp" onClick="go_cmdurl('更新数据',this)" target="content3" title="更新数据"><span>更新数据</span></a></li>
											<li id="now1c"><a href="user_setting.asp" onClick="go_cmdurl('综合设置',this)" target="content3" title="综合设置"><span>综合设置</span></a></li>
<%
Dim rstm
Set rstm=oblog.Execute("select top 1 userid From oblog_admin Where userid=" & oblog.l_uid)
If Not rstm.Eof Then
%>
											<li id="now1d"><a href="<%=SYSFOLDER_MANAGER%>/m_index.asp"  target="_blank" title="进入内容管理操作界面"><span><font color="red">内容管理员</font></span></a></li>
<%
Set rstm=Nothing
End If
%>
										</ul>
										<ul id="dleft_tab2" style="display:none;">
											<li id="now21"><a href="user_post.asp" onClick="go_cmdurl('发布日志',this)" target="content3" title="发布日志"><span>发布日志</span></a></li>
											<li id="now22"><a href="user_blogmanage.asp" onClick="go_cmdurl('日志管理',this)" target="content3" title="日志管理"><span>日志管理</span></a></li>
											<li id="now23"><a href="user_blogmanage.asp?usersearch=5" onClick="go_cmdurl('草稿箱',this)" target="content3" title="草稿箱"><span>草稿箱<span id="sdraft_num"> </span></span></a></li>
											<li id="now24"><a href="user_blogmanage.asp?usersearch=6" onClick="go_cmdurl('回收站',this)" target="content3" title="回收站"><span>回收站<span id="del_num"></span></span></a></li>
											<li id="now25"><a href="user_subject.asp" onClick="go_cmdurl('日志专题',this)" target="content3" title="日志专题"><span>日志专题</span></a></li>
											<li id="now26"><a href="user_blogmanage.asp?action=downlog" onClick="go_cmdurl('备份日志',this)" target="content3" title="备份日志"><span>备份日志</span></a></li>
											<li id="now27"><a href="user_comments.asp" onClick="go_cmdurl('评论管理',this)" target="content3" title="评论管理"><span>评论管理</span></a></li>
											<li id="now28"><a href="user_messages.asp" onClick="go_cmdurl('留言管理',this)" target="content3" title="留言管理"><span>留言管理</span></a></li>
											<li id="now29"><a href="user_tb.asp" onClick="go_cmdurl('引用通告',this)" target="content3" title="引用通告"><span>引用通告</span></a></li>
										</ul>
										<ul id="dleft_tab3" style="display:none;">
											<li id="now31"><a href="user_post.asp?t=1&tMode=normal" onClick="go_cmdurl('发布相片',this)" target="content3" title="发布相片"><span>发布相片</span></a></li>
											<li id="now32"><a href="user_post.asp?t=1&action=showphoto" onClick="go_cmdurl('我的相册',this)" target="content3" title="浏览相册"><span>浏览相册</span></a></li>
											<li id="now33"><a href="user_photo.asp" onClick="go_cmdurl('相片管理',this)" target="content3" title="相片管理"><span>相片管理</span></a></li>
											<li id="now34"><a href="user_Albumcomments.asp" onClick="go_cmdurl('相册评论',this);" target="content3" title="相册评论"><span>相册评论</span></a></li>
											<li id="now35"><a href="user_subject.asp?t=1" onClick="go_cmdurl('相册分类',this)" target="content3" title="相册分类"><span>相册分类</span></a></li>
											<li id="now36"><a href="user_post.asp?t=2" onClick="go_cmdurl('大头贴',this)" target="content3" title="大头贴"><span>大头贴</span></a></li>
										</ul>
										<ul id="dleft_tab4" style="display:none;">
											<li id="now41"><a href="user_team.asp" onClick="go_cmdurl('<%=oblog.CacheConfig(69)%>最近话题',this)" target="content3" title="<%=oblog.CacheConfig(69)%>最近话题"><span><%=oblog.CacheConfig(69)%>最近话题</span></a></li>
											<li id="now42"><a href="user_team.asp?action=listmanageteam" onClick="go_cmdurl('我管理的<%=oblog.CacheConfig(69)%>',this)" target="content3" title="我管理的<%=oblog.CacheConfig(69)%>"><span>我管理的<%=oblog.CacheConfig(69)%></span></a></li>
											<li id="now43"><a href="user_team.asp?action=listjoinedteam" onClick="go_cmdurl('我加入的<%=oblog.CacheConfig(69)%>',this)" target="content3" title="我加入的<%=oblog.CacheConfig(69)%>"><span>我加入的<%=oblog.CacheConfig(69)%></span></a></li>
											<li id="now44"><a href="user_team.asp?action=creatteam" onClick="go_cmdurl('创建新<%=oblog.CacheConfig(69)%>',this)" target="content3" title="创建新<%=oblog.CacheConfig(69)%>"><span>创建新<%=oblog.CacheConfig(69)%></span></a></li>
											<li id="now45"><a href="user_team.asp?action=members&cmd=1" onClick="go_cmdurl('发出的邀请',this)" target="content3" title="发出的邀请"><span>发出的邀请</span></a></li>
											<li id="now46"><a href="user_team.asp?action=members&cmd=2" onClick="go_cmdurl('收到的邀请',this)" target="content3" title="收到的邀请"><span>收到的邀请</span></a></li>
											<li id="now47"><a href="user_team.asp?action=members&cmd=3" onClick="go_cmdurl('发出的申请',this)" target="content3" title="发出的申请"><span>发出的申请</span></a></li>
											<li id="now48"><a href="user_team.asp?action=members&cmd=4" onClick="go_cmdurl('收到的申请',this)" target="content3" title="收到的申请"><span>收到的申请</span></a></li>
										</ul>
										<ul id="dleft_tab5" style="display:none;">
											<li id="now51"><a href="user_files.asp" onClick="go_cmdurl('所有文件',this)" target="content3" title="所有文件"><span>所有文件</span></a></li>
											<li id="now52"><a href="user_files.asp?cmd=1" onClick="go_cmdurl('图片文件',this)" target="content3" title="图片文件"><span>图片文件</span></a></li>
											<li id="now53"><a href="user_files.asp?cmd=2" onClick="go_cmdurl('压缩文件',this)" target="content3" title="FLASH"><span>ＦＬＡＳＨ</span></a></li>
											<li id="now54"><a href="user_files.asp?cmd=3" onClick="go_cmdurl('文档文件',this)" target="content3" title="音频文件"><span>音频文件</span></a></li>
											<li id="now55"><a href="user_files.asp?cmd=4" onClick="go_cmdurl('文档文件',this)" target="content3" title="视频文件"><span>视频文件</span></a></li>
											<li id="now56"><a href="user_files.asp?cmd=5" onClick="go_cmdurl('文档文件',this)" target="content3" title="压缩文件"><span>压缩文件</span></a></li>
											<li id="now57"><a href="user_files.asp?cmd=6" onClick="go_cmdurl('文档文件',this)" target="content3" title="文档文件"><span>文档文件</span></a></li>
											<li id="now58"><a href="user_files.asp?cmd=999" onClick="go_cmdurl('文档文件',this)" target="content3" title="文档文件"><span>其他文件</span></a></li>
										</ul>
										<ul id="dleft_tab6" style="display:none;">
											<li id="now61"><a href="user_template.asp?action=showconfig" onClick="go_cmdurl('选择模板',this)" target="content3" title="选择模板"><span>选择模板</span></a></li>
											<li id="now62"><a href="user_template.asp?action=modiconfig&editm=1" onClick="go_cmdurl('改主模板',this)" target="content3" title="改主模板"><span>改主模板</span></a></li>
											<li id="now63"><a href="user_template.asp?action=modiviceconfig&editm=1" onClick="go_cmdurl('改副模板',this)" target="content3" title="改副模板"><span>改副模板</span></a></li>
											<li id="now64"><a href="user_template.asp?action=bakskin" onClick="go_cmdurl('备份模板',this)" target="content3" title="备份模板"><span>备份模板</span></a></li>
										</ul>
										<ul id="dleft_tab8" style="display:none;">
<%If oblog.CacheConfig(81) = "1" Then %>
											<li id="now81"><a href="#nogo" onClick="pmusicurl();" title="AOBO音乐盒"><span>AOBO音乐盒</span></a></li>
<%End if%>
										</ul>
										<ul id="dleft_tab7" style="display:none;">
										<li><img src="images/loading.gif">正在加载...</li>
										</ul>
									</div>
									<div class="clear"></div>
								</div>
							</td>
						</tr>
						<tr class="t3">
							<td><div class="menu_end"></div></td>
						</tr>
					</table>
				</div>
			</td>
			<td class="t2">
				<div id="cnt">
					<div id="dTab1">
						<iframe src="user_index_frame1.asp" name="content1" frameborder="0" scrolling="no"></iframe>
					</div>
					<div id="dTab2">
						<iframe <%If oblog.CacheConfig(81) = "1" Then %><%End If %> name="content2" frameborder="0" scrolling="no"></iframe>
					</div>
					<div id="dTab3">
						<iframe src="<%=MainUrl%>"  name="content3" id="content3" frameborder="0" scrolling="no"></iframe>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/cnt.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
</body>
</html>
<script>
//读订阅列表
function getfeedlist(){
	document.getElementById("dleft_tab7").innerHTML="<span style='margin:20px;'>正在加载...</span>";
	var Ajax = new oAjax("AjaxServer.asp",showfeedlist);
	var arrKey = new Array("action");
	var arrValue = new Array('getfeedlist');
	Ajax.Post(arrKey,arrValue);
}
function showfeedlist(arrobj){
	if (arrobj){
		document.getElementById("dleft_tab7").innerHTML=arrobj[0];
	}
}

function su_click(obj){
	if(obj.className == 'open')
	{obj.className = 'close';}
	else{obj.className = 'open';}

}
//修改tab3标题
function show_title(str){
	document.getElementById("dnow99").innerHTML=str;
	//document.getElementById("dnow99").style.display='block';
}

//读短消息
function getpm(){
	//var rsslist = new Get_rsslist('AjaxServer.asp?action=getfeedlist');
	var Ajax = new oAjax("AjaxServer.asp",showpm);
	var arrKey = new Array("action","username");
	var arrValue = new Array('getpm',"<%=oblog.l_uname%>");
	Ajax.Post(arrKey,arrValue);
}
function showpm(arrobj){
	if (arrobj){
		document.getElementById("ob_pm").innerHTML=arrobj[0];
		<%If En_SMS_Sound=1 Then%>
		if (arrobj[0]!="(0)"){
			document.getElementById("ob_pm").innerHTML=document.getElementById("ob_pm").innerHTML+"<EMBED SRC='oblogstyle/newsms.wav' HIDDEN=true AUTOSTART=true LOOP=false>";
		}
		<%End If%>
	}
}

//读草稿数
function get_draft(){
	//var rsslist = new Get_rsslist('AjaxServer.asp?action=getfeedlist');
	var Ajax = new oAjax("AjaxServer.asp",show_draft);
	var arrKey = new Array("action","userid");
	var arrValue = new Array('get_draft',"<%=oblog.l_uid%>");
	Ajax.Post(arrKey,arrValue);
}
function show_draft(arrobj){
	if (arrobj){
		document.getElementById("sdraft_num").innerHTML=arrobj[0];
		document.getElementById("del_num").innerHTML=arrobj[1];
	}
}

function go_cmdurl(title,tabid){
	show_title(title);
	switchTab('TabPage1','Tab3');
	menu(document.getElementById('Tab3'));
	dleft_tab_active('TabPage3',tabid);
}

function u_init(){
	<%=jscmd%>
	getpm();
	get_draft();
	setInterval(getpm,<%=oblog.CacheConfig(8)%>*60000);
}
u_init();
</script>
<%Set oblog = Nothing%>