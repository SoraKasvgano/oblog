<!--#include file="inc/inc_sys.asp"-->
<%
select case Request("action")
case "top"
	call admin_top()
case "left"
	call m_left()
case "main"
	call m_main()
case "state"
	If Application(cache_name_user&"_systemstate")<>"stop" Then
		Application(cache_name_user&"_systemstate")="stop"
	Else
		Application(cache_name_user&"_systemstate")="run"
	End If
	Application(cache_name_user&"_systemnote")=Request.Form("systemnote")
	Response.Write "<script language=javascript>parent.location.href=""m_index.asp"";</script>"
case else
	call main()
end select

sub main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/admin/style.css" type="text/css" />
</head>
<frameset rows="*" cols="180,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
  <frame name="left" scrolling="auto" marginwidth="0" marginheight="0" src="m_index.asp?action=left">
  <frameset framespacing="0" border="false" rows="20,*" frameborder="0" scrolling="yes">
    <frame name="top" scrolling="no" src="m_index.asp?action=top">
    <frame name="main" scrolling="auto" src="m_index.asp?action=main">
  </frameset>
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>你的浏览器版本过低！！！本系统要求IE5及以上版本才能使用本系统。</p>
  </body>
</noframes>
</html>
<%
end sub

sub admin_top()
%>
<html>
<head>
<title>oBlog后台管理页面</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/style.css" type="text/css" />
<style type="text/css">
a:link {
	color:#000000;
	text-decoration:none;
	font-size: 12px;
}
a:hover {color:#CC3300;}
a:visited {color:#000000;text-decoration:none}

td {FONT-SIZE: 9pt;COLOR: #000000; FONT-FAMILY: "宋体"}
img {filter:Alpha(opacity:100); chroma(color=#FFFFFF)}
</style>
<base target="main">
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" height="100%" border=0 cellpadding=0 cellspacing=0>
  <tr valign=middle>
    <td width=10></td>
	<td width=50><a href="m_pwd.asp">修改密码</a></td>
    <td align="left" width="500"><div id="ob4news"></div></td>
    <td width="50" align="left"><a href="../index.asp" target="_blank">站点首页</a></td>
  </tr>
</table>
</body>
</html>
<script src="http://www.oblog.cn/oblog4news.asp"></script>
<%end sub
sub m_left()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/admin/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body style="overflow-x:hidden;text-align:left;">
<div id="logo"></div>
<div id="TabPage1">
<!--TabPage1-->
	<div class="left_top"></div>
	<ul class="left_conten">
		<li><a href="m_index.asp?action=main" target="main"><strong>管理首页</strong></a>|<a href="m_login.asp?action=logout" target="_top"><strong>退出</strong></a></li>
		<li><a>用户名：<%=m_name%></a></li>
		<li><a>权　限：<%
		Dim trs
		Set trs = oblog.Execute ("select r_name FROM oblog_roles WHERE roleid = " &OB_IIF(Session("roleid"),0))
		If Not trs.EOF Then
			Response.Write trs(0)
		Else
			If session("AdminName") <> "" Then
				Response.Write  "系统管理员"
			Else
				Response.Write "内容管理员"
			End if
		End If
		trs.close
		Set trs = Nothing
		%></a></li>
	</ul>
	<div <%=CheckDisplay(1)%>>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_1)" >
		<li class="left_top_left left">常规管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_1">
	<%If CheckAccess("r_words") Then%>
		<li><a href="m_words.asp" target="main" id="Tab101" onClick="javascript:switchTab('TabPage1','Tab101');">关键字管理</a></li>
	<%End if%>
	<%If CheckAccess("r_IP") Then%>
		<li><a href="m_lockip.asp" target="main" id="Tab102" onClick="javascript:switchTab('TabPage1','Tab102');">限制IP管理</a></li>
	<%End if%>
	<%If CheckAccess("r_site_news") Then%>
		<li><a href="../admin_edit.asp?do=4" target="main" id="Tab103" onClick="javascript:switchTab('TabPage1','Tab103');">用户后台通知</a></li>
	<%End if%>
	<%If CheckAccess("r_user_news") Then%>
		<li><a href="m_pmall.asp" target="main" id="Tab104" onClick="javascript:switchTab('TabPage1','Tab104');">发送站内短信</a></li>
	<%End if%>
	<%If CheckAccess("r_site_count") Then%>
		<li><a href="m_count.asp" target="main" id="Tab105" onClick="javascript:switchTab('TabPage1','Tab105');">更新系统数据</a></li>
	<%End if%>
	</ul>
	</div>
	<div <%=CheckDisplay(2)%>>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_2)" >
		<li class="left_top_left left">内容管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_2">
	<%If CheckAccess("r_user_blog") Then%>
		<li><a href="m_blog.asp" target="main" id="Tab201" onClick="javascript:switchTab('TabPage1','Tab201');">日志管理</a></li>
	<%End if%>
	<%If CheckAccess("r_user_rblog") Then%>
		<li><a href="m_r_blog.asp" target="main" id="Tab202" onClick="javascript:switchTab('TabPage1','Tab202');">回收站管理</a></li>
	<%End if%>
	<%If CheckAccess("r_user_blog") Then%>
		<li><a href="m_blog.asp?cmd=3" target="main" id="Tab203a" onClick="javascript:switchTab('TabPage1','Tab203a');">精华日志</a>|<a href="m_blog.asp?cmd=2" target="main" id="Tab203b" onClick="javascript:switchTab('TabPage1','Tab203b');">待审日志</a></li>
	<%End if%>
	<%If CheckAccess("r_user_cmt") Then%>
		<li><a href="m_comments.asp" target="main" id="Tab204a" onClick="javascript:switchTab('TabPage1','Tab204a');">评论管理</a>|<a href="m_comments.asp?cmd=1" target="main" id="Tab204b" onClick="javascript:switchTab('TabPage1','Tab204b');">待审评论</a></li>
	<%End if%>
	<%If CheckAccess("r_user_cmt") Then%>
		<li><a href="m_messages.asp" target="main" id="Tab205a" onClick="javascript:switchTab('TabPage1','Tab205a');">留言管理</a>|<a href="m_messages.asp?cmd=1" target="main" id="Tab205b" onClick="javascript:switchTab('TabPage1','Tab205b');">待审留言</a></li>
	<%End if%>
	<%If CheckAccess("r_user_tag") Then%>
		<li><a href="m_tags.asp" target="main" id="Tab206" onClick="javascript:switchTab('TabPage1','Tab206');">TAG审核</a></li>
	<%End if%>
	<%If CheckAccess("r_album_comment") Then%>
		<li><a href="m_album_comments.asp" target="main" id="Tab207a" onClick="javascript:switchTab('TabPage1','Tab207a');">相册评论</a>|<a href="m_album_comments.asp?cmd=1" target="main" id="Tab207b" onClick="javascript:switchTab('TabPage1','Tab207b');">待审评论</a></li>
	<%End if%>
	<%If CheckAccess("r_user_digg") Then%>
		<li><a href="m_userdigg.asp" target="main" id="Tab208a" onClick="javascript:switchTab('TabPage1','Tab208a');">DIGG管理</a>|<a href="m_digg.asp" target="main" id="Tab208b" onClick="javascript:switchTab('TabPage1','Tab208b');">DIGG记录</a></li>
		<li><a href="m_digg.asp?cmd=1" target="main" id="Tab209" onClick="javascript:switchTab('TabPage1','Tab209');">反映问题管理</a></li>
	<%End if%>
	</ul>
	</div>
	<div <%=CheckDisplay(3)%>>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_3)" >
		<li class="left_top_left left"><%=oblog.CacheConfig(69)%>管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_3">
	<%If CheckAccess("r_group_user") Then%>
		<li><a href="m_team.asp?cmd=2" target="main" id="Tab301" onClick="javascript:switchTab('TabPage1','Tab301');"><%=oblog.CacheConfig(69)%>管理</a></li>
		<li><a href="m_team.asp" target="main" id="Tab302" onClick="javascript:switchTab('TabPage1','Tab302');">待审<%=oblog.CacheConfig(69)%></a></li>
	<%End if%>
	<%If CheckAccess("r_group_blog") Then%>
		<li><a href="m_post.asp" target="main" id="Tab303" onClick="javascript:switchTab('TabPage1','Tab303');"><%=oblog.CacheConfig(69)%>内容管理</a></li>
	<%End if%>
	</ul>
	</div>
	<div <%=CheckDisplay(4)%>>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_4)" >
		<li class="left_top_left left">上传文件管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_4">
	<%If CheckAccess("r_user_upfiles") Then%>
		<li><a href="m_uploadfile_user.asp" target="main" id="Tab401" onClick="javascript:switchTab('TabPage1','Tab401');">上传管理用户清单</a></li>
	<%End if%>
	<%If CheckAccess("r_list_upfiles") Then%>
		<li><a href="m_uploadfile.asp" target="main" id="Tab402" onClick="javascript:switchTab('TabPage1','Tab402');">上传管理文件清单</a></li>
	<%End if%>
	</ul>
	</div>
	<div <%=CheckDisplay(5)%>>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_5)">
		<li class="left_top_left left">用户管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_5">
	<%If CheckAccess("r_user_all") Then%>
		<li><a href="m_user.asp" target="main" id="Tab501" onClick="javascript:switchTab('TabPage1','Tab501');">全部用户管理</a></li>
		<li><a href="m_user.asp?cmd=6" target="main" id="Tab503" onClick="javascript:switchTab('TabPage1','Tab503');">待审核用户列表</a></li>
		<li><a href="m_user.asp?cmd=9" target="main" id="Tab504" onClick="javascript:switchTab('TabPage1','Tab504');">敏感用户列表</a></li>
		<li><a href="m_user.asp?Action=Update" target="main" id="Tab106" onClick="javascript:switchTab('TabPage1','Tab106');">生成用户静态页</a></li>
		<%If CheckAccess("r_user_Admin") Then%>
		<li><a href="m_user.asp?action=gouser1" target="main" id="Tab506" onClick="javascript:switchTab('TabPage1','Tab506');">进入用户管理界面</a></li>
		<%End If%>
	<%End if%>
	<%If CheckAccess("r_blogstar") Then%>
		<li><a href="m_blogstar.asp" target="main" id="Tab502" onClick="javascript:switchTab('TabPage1','Tab502');">博客之星管理</a></li>
	<%End if%>
<!-- 		<li><a href="m_user.asp?cmd=10" target="main" id="Tab505" onClick="javascript:switchTab('TabPage1','Tab505');">系统自动锁定用户</a></li> -->
		<%If CheckAccess("r_user_name") Then%>
		<li><a href="m_rename.asp" target="main" id="Tab507" onClick="javascript:switchTab('TabPage1','Tab507');">用户改名</a></li>
		<%End If%>
	<%If CheckAccess("r_user_add") Then%>
		<li><a href="../reg.asp" target="main" id="Tab508" onClick="javascript:switchTab('TabPage1','Tab508');" title="系统关闭注册时，管理员仍然可以通过此方式添加新的用户">增加新用户</a></li>
	<%End if%>
	</ul>
	</div>
	<div <%=CheckDisplay(6)%>>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_6)" >
		<li class="left_top_left left">模板管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_6">
	<%If CheckAccess("r_skin_sys") Then%>
		<li><a href="m_sysskin.asp?action=addskin" target="main" id="Tab601a" onClick="javascript:switchTab('TabPage1','Tab601a');">添加系统模板</a>|<a href="m_sysskin.asp?action=showskin" target="main" id="Tab601b" onClick="javascript:switchTab('TabPage1','Tab601b');">管理</a></li>
	<%End If%>
	<%If CheckAccess("r_skin_user") Then%>
		<li><a href="m_userskin.asp?action=addskin" target="main" id="Tab602a" onClick="javascript:switchTab('TabPage1','Tab602a');">添加用户模板</a>|<a href="m_userskin.asp?action=showskin&ispass=1" target="main" id="Tab602b" onClick="javascript:switchTab('TabPage1','Tab602b');">管理</a></li>
	<%End If%>
	</ul>
	</div>
	<div class="left_end"></div>
<!--/TabPage1-->
</div>
<div id="cnt"></div>
</body>
</html>
<%
Response.Write("<script src=""http://www.oblog.cn/count/count.asp?a="&oblog.cacheconfig(3)&"&b="&oblog.cacheconfig(4)&"&c="&oblog.setup(1,0)&"&d="&oblog.ver&"&e="&is_sqldata&"&f="&oblog.cacheConfig(11)&"&g="&oblog.setup(4,0)&"""></script>")
end sub
sub m_main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">oBlog后台管理首页</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 class="topbg"><strong>oBlog 帮 助</strong>
  <tr>
    <td height=23 class="tdbg">1、<strong>将用户前台屏蔽以后此用户发布的所有文章(包括照片)都不会在首页被调用</strong>。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">2、<a href="m_skin_help.asp" target="_blank"><strong>系统模版及用户模版的标记说明请点击这里</strong>。</a></td>
  </tr>
  <tr>
    <td height=23 class="tdbg"><p>3、用户权限：后台管理员可以建立不同的用户组，赋予不同的权限。</p>
    </td>
  </tr>
  <tr>
    <td class="tdbg">4、将用户锁定以后，此用户的blog页面也将被屏蔽。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">5、将IP屏蔽以后，此IP用户将不能登录，且不能发表评论及留言。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">6、将博客设置为推荐，必须在后台修改用户资料才能实现。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">7、若上传文件不正常，请检查是否文件尺寸过大及服务器是否支持fso。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">8、有任何问题，请咨询oBlog官方网站<a href="http://www.oBlog.cn" target="_blank">http://www.oBlog.cn</a>。</td>
  </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</body>
</html>
<%
end Sub
Set oblog = Nothing
%>