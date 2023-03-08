<!--#include file="inc/inc_sys.asp"-->
<%
select case Request("action")
case "top"
	call admin_top()
case "left"
	call admin_left()
case "main"
	call admin_main()
case "state"
	If Application(cache_name_user&"_systemstate")<>"stop" Then
		Application(cache_name_user&"_systemstate")="stop"
		EventLog "临时关闭了本站",oblog.NowUrl&"?"&Request.QueryString
	Else
		Application(cache_name_user&"_systemstate")="run"
		EventLog "开启了本站",oblog.NowUrl&"?"&Request.QueryString
	End If
	Application(cache_name_user&"_systemnote")=Request.Form("systemnote")
	Response.Write "<script language=javascript>parent.location.href=""admin_index.asp"";</script>"
case "enmod"
	Dim enStr
	enStr=OB_IIF(Request("encomment"),"0")
	enStr=enStr & "," & OB_IIF(Request("enargue"),"0")
	enStr=enStr & "," & OB_IIF(Request("enblog"),"0")
	enStr=enStr & "," & OB_IIF(Request("entb"),"0")
	Application(cache_name_user&"_systemenmod")=enStr
	EventLog "临时关闭（重新开启）了本站的部分功能",oblog.NowUrl&"?"&Request.QueryString
	Response.Write "<script language=javascript>parent.location.href=""admin_index.asp"";</script>"
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
  <frame name="left" scrolling="auto" marginwidth="0" marginheight="0" src="admin_index.asp?action=left">
  <frameset framespacing="0" border="false" rows="20,*" frameborder="0" scrolling="yes">
    <frame name="top" scrolling="no" src="admin_index.asp?action=top">
    <frame name="main" scrolling="auto" src="admin_index.asp?action=main">
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
	<td width=50><a href="admin_adminmodifypwd.asp">修改密码</a></td>
     <td align="left" width="500"><span id="ob4news"></span></td>
    <td width="50" align="left"><a href="../index.asp" target="_blank">站点首页</a></td>
  </tr>
</table>
</body>
</html>
<script src="http://www.oblog.cn/oblog4news.asp"></script>
<%end sub
sub admin_left()
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
	<ul class="left_conten" id="oblog_0">
		<li><a href="admin_index.asp?action=main" target="main"><strong>管理首页</strong></a>|<a href="admin_login.asp?action=logout"  target=_top><strong>退出</strong></a></li>
		<li><a>用户名：<%= session("adminname") %></a></li>
		<li><a href="../<%=SYSFOLDER_MANAGER%>/m_index.asp" target=_blank>进入内容管理</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_1)">
		<li class="left_top_left left">系统设置</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_1">
		<li><a href="admin_setup.asp" target="main" id="Tab101" onClick="javascript:switchTab('TabPage1','Tab101');">网站信息配置</a></li>
		<li><a href="admin_com.asp" target="main" id="Tab102" onClick="javascript:switchTab('TabPage1','Tab102');">服务器组件配置</a></li>
		<li><a href="admin_syslog.asp" target="main" id="Tab405" onClick="javascript:switchTab('TabPage1','Tab405');">系统操作日志管理</a></li>
		<li><a href="admin_userclass.asp" target="main" id="Tab103" onClick="javascript:switchTab('TabPage1','Tab103');">系统博客分类管理</a></li>
		<li><a href="admin_logclass.asp" target="main" id="Tab104" onClick="javascript:switchTab('TabPage1','Tab104');">系统日志分类管理</a></li>
		<li><a href="admin_logclass.asp?t=1" target="main" id="Tab105" onClick="javascript:switchTab('TabPage1','Tab105');">系统相册分类管理</a></li>
		<li><a href="admin_logclass.asp?t=2" target="main" id="Tab107" onClick="javascript:switchTab('TabPage1','Tab107');">系统群组分类管理</a></li>
		<li><a href="../admin_edit.asp?do=3" target="main" id="Tab106a" onClick="javascript:switchTab('TabPage1','Tab106a');">修改注册条款</a>|<a href="admin_note.asp?action=do3" target="main" id="Tab106b" onClick="javascript:switchTab('TabPage1','Tab106b');">文本</a></li>
		<li><a href="admin_js.asp" target="main" id="Tab108" onClick="javascript:switchTab('TabPage1','Tab108');">JS调用管理</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_2)">
		<li class="left_top_left left">常规设置</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_2">
		<li><a href="admin_score.asp" target="main" id="Tab201" onClick="javascript:switchTab('TabPage1','Tab201');">网站积分制度</a></li>
		<li><a href="admin_ask.asp" target="main" id="Tab209" onClick="javascript:switchTab('TabPage1','Tab209');">自定义验证问题管理</a></li>
		<li><a href="../admin_edit.asp?do=1" target="main" id="Tab202a" onClick="javascript:switchTab('TabPage1','Tab202a');">修改友情链接</a>|<a href="admin_note.asp?action=do1" target="main" id="Tab202b" onClick="javascript:switchTab('TabPage1','Tab202b');">文本</a></li>
		<li><a href="../admin_edit.asp?do=2" target="main" id="Tab203a" onClick="javascript:switchTab('TabPage1','Tab203a');">修改网站公告</a>|<a href="admin_note.asp?action=do2" target="main" id="Tab203b" onClick="javascript:switchTab('TabPage1','Tab203b');">文本</a></li>
		<li><a href="../admin_edit.asp?do=4" target="main" id="Tab204a" onClick="javascript:switchTab('TabPage1','Tab204a');">用户后台通知</a>|<a href="admin_note.asp?action=do4" target="main" id="Tab204b" onClick="javascript:switchTab('TabPage1','Tab204b');">文本</a></li>
		<li><a href="admin_lockip.asp" target="main" id="Tab205" onClick="javascript:switchTab('TabPage1','Tab205');">限制IP管理</a></li>
		<li><a href="admin_report.asp" target="main" id="Tab2055" onClick="javascript:switchTab('TabPage1','Tab2055');">反映问题管理</a></li>
		<li><a href="admin_count.asp" target="main" id="Tab206" onClick="javascript:switchTab('TabPage1','Tab206');">更新系统数据</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_3)">
		<li class="left_top_left left">广告管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_3">
		<li><a href="admin_ad.asp" target="main" id="Tab301" onClick="javascript:switchTab('TabPage1','Tab301');">用户页面广告管理</a></li>
		<li><a href="admin_teamad.asp" target="main" id="Tab302" onClick="javascript:switchTab('TabPage1','Tab302');">群组页面广告管理</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_4)">
		<li class="left_top_left left">管理员级别及设置</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_4">
		<li><a href="admin_admin.asp?Action=Add" target="main" id="Tab401" onClick="javascript:switchTab('TabPage1','Tab401');">添加新的管理员</a></li>
		<li><a href="admin_role.asp?action=add" target="main" id="Tab402" onClick="javascript:switchTab('TabPage1','Tab402');">内容管理员分级</a></li>
		<li><a href="admin_role.asp" target="main" id="Tab403" onClick="javascript:switchTab('TabPage1','Tab403');">内容管理员列表</a></li>
		<li><a href="admin_admin.asp" target="main" id="Tab404" onClick="javascript:switchTab('TabPage1','Tab404');">全部管理员列表</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_5)">
		<li class="left_top_left left">用户等级及设置</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_5">
		<li><a href="admin_user.asp" target="main" id="Tab506" onClick="javascript:switchTab('TabPage1','Tab506');">全部用户管理</a></li>
		<li><a href="admin_groups.asp?action=add" target="main" id="Tab501" onClick="javascript:switchTab('TabPage1','Tab501');">新增用户等级</a></li>
		<li><a href="admin_groups.asp" target="main" id="Tab502" onClick="javascript:switchTab('TabPage1','Tab502');">管理用户等级</a></li>
		<li><a href="admin_rename.asp" target="main" id="Tab503" onClick="javascript:switchTab('TabPage1','Tab503');">用户改名</a></li>
		<li><a href="admin_userdir.asp" target="main" id="Tab504" onClick="javascript:switchTab('TabPage1','Tab504');">用户目录管理</a></li>
		<li><a href="admin_user.asp?Action=Update" target="main" id="Tab505" onClick="javascript:switchTab('TabPage1','Tab505');">生成用户静态页</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_6)">
		<li class="left_top_left left">模板管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_6">
		<li><a href="admin_sysskin.asp?action=addskin" target="main" id="Tab601a" onClick="javascript:switchTab('TabPage1','Tab601a');">添加系统模板</a>|<a href="admin_sysskin.asp?action=showskin" target="main" id="Tab601b" onClick="javascript:switchTab('TabPage1','Tab601b');">管理</a></li>
		<li><a href="admin_skin.asp?action=insys1" target="main" id="Tab602a" onClick="javascript:switchTab('TabPage1','Tab602a');">系统模板导入</a>|<a href="admin_skin.asp?action=outsys" target="main" id="Tab602b" onClick="javascript:switchTab('TabPage1','Tab602b');">导出</a></li>
		<li><a href="admin_userskin.asp?action=skinclass" target="main" id="Tab603" onClick="javascript:switchTab('TabPage1','Tab603');">用户模板分类管理</a></li>
		<li><a href="admin_userskin.asp?action=addskin" target="main" id="Tab604a" onClick="javascript:switchTab('TabPage1','Tab604a');">添加用户模板</a>|<a href="admin_userskin.asp?action=showskin&ispass=1" target="main" id="Tab604b" onClick="javascript:switchTab('TabPage1','Tab604b');">管理</a></li>
		<li><a href="admin_skin.asp?action=inuser1" target="main" id="Tab605a" onClick="javascript:switchTab('TabPage1','Tab605a');">用户模板导入</a>|<a href="admin_skin.asp?action=outuser" target="main" id="Tab605b" onClick="javascript:switchTab('TabPage1','Tab605b');">导出</a></li>
		<li><a href="admin_teamskin.asp?action=skinclass" target="main" id="Tab608" onClick="javascript:switchTab('TabPage1','Tab608');">群组模板分类管理</a></li>
		<li><a href="admin_teamskin.asp?action=addskin" target="main" id="Tab606a" onClick="javascript:switchTab('TabPage1','Tab606a');">添加群组模板</a>|<a href="admin_teamskin.asp?action=showskin&ispass=1" target="main" id="Tab606b" onClick="javascript:switchTab('TabPage1','Tab606b');">管理</a></li>
		<li><a href="admin_skin.asp?action=inteam1" target="main" id="Tab607a" onClick="javascript:switchTab('TabPage1','Tab607a');">群组模板导入</a>|<a href="admin_skin.asp?action=outteam" target="main" id="Tab607b" onClick="javascript:switchTab('TabPage1','Tab607b');">导出</a></li>
	</ul>
	<div class="left_end"></div>
<%
'如果是SQL Server数据库则不再显示该节
If IS_SQLDATA=0 Then
%>
	<ul class="left_top" onClick="menu(oblog_7)">
		<li class="left_top_left left">数据库管理</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_7">
		<li><a href="admin_database.asp?Action=Compact" target="main" id="Tab701" onClick="javascript:switchTab('TabPage1','Tab701');">压缩数据库</a></li>
		<li><a href="admin_database.asp?Action=Backup" target="main" id="Tab702" onClick="javascript:switchTab('TabPage1','Tab702');">备份数据库</a></li>
		<li><a href="admin_database.asp?Action=Restore" target="main" id="Tab703" onClick="javascript:switchTab('TabPage1','Tab703');">恢复数据库</a></li>
		<li><a href="admin_database.asp?Action=SpaceSize" target="main" id="Tab704" onClick="javascript:switchTab('TabPage1','Tab704');">系统空间占用</a></li>
	</ul>
	<div class="left_end"></div>
<%End If%>
	<ul class="left_top" onClick="menu(oblog_8)">
		<li class="left_top_left left">系统分析</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_8">
		<li><a href="admin_analyze.asp?action=year" target="main" id="Tab801" onClick="javascript:switchTab('TabPage1','Tab801');">数据年度分析</a></li>
		<li><a href="admin_analyze.asp?action=month" target="main" id="Tab802" onClick="javascript:switchTab('TabPage1','Tab802');">数据月度分析</a></li>
		<li><a href="admin_analyze.asp?action=hour" target="main" id="Tab803" onClick="javascript:switchTab('TabPage1','Tab803');">数据时段分析</a></li>
	</ul>
	<div class="left_end"></div>
<!--/TabPage1-->
</div>
<div id="cnt"></div>
	<ul class="left_top" onClick="menu(oblog_9)">
		<li class="left_top_left left">oBlog信息</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_9">
		<li>版权所有:<a>北京傲博致远计算机技术有限公司</a></li>
		<li>网站地址:<a href="http://www.oBlog.cn" target="_blank">oBlog.cn</a></li>
		<li>技术支持:<a href="http://bbs.oBlog.cn" target="_blank">oBlog论坛</a></li>
	</ul>
	<div class="left_end"></div>
<!--/TabPage1-->
</div>
<div id="cnt"></div>
</body>
</html>
<%
Response.Write("<script src=""http://www.oblog.cn/count/count.asp?a="&oblog.cacheconfig(3)&"&b="&oblog.cacheconfig(4)&"&c="&oblog.setup(1,0)&"&d="&oblog.ver&"&e="&is_sqldata&"&f="&oblog.cacheConfig(11)&"&g="&oblog.setup(4,0)&"""></script>")
end sub
sub admin_main()
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
<table width="100%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
<td height=25 colspan=2 class="topbg"><strong>系 统 状 态</strong></th></tr>
     <tr>
    <td width="20%" class="tdbg" height=23>当前系统版本</td>
    <td width="80%" class="tdbg">&nbsp;&nbsp;<font color=blue><%=ver%></font></td>
  </tr>
     <tr>
    <td width="20%" class="tdbg" height=23>最新系统版本</td>
    <td width="80%" class="tdbg">&nbsp;&nbsp;<span id="ob4version"></span></td>
  </tr>

     <tr>
    <td width="20%" class="tdbg" height=23>当前系统状态</td>
    <td width="80%" class="tdbg">&nbsp;&nbsp;<font color=red>
    <%
    	Dim strSubmit
    	If Application(cache_name_user&"_systemstate")="stop" Then
    		Response.Write "关闭中"
    		strSubmit="重新启动"
    	Else
    		Response.Write "正常运行中"
    		strSubmit="关闭系统"
    	End If
    %></font>
    </td>
  </tr>
   <form name="systemcontrol" method="post" action="admin_index.asp?action=state">
	  <tr>
	    <td width="20%" class="tdbg" height=23>关闭时提示：</td>
	    <td width="80%" class="tdbg">
	          	<textarea cols=60 rows=5 name="systemnote"><%
	          		If Application(cache_name_user&"_systemnote")="" Then
	          			Response.Write "系统正在维护，稍后开放"
	          		Else
	          			Response.Write Application(cache_name_user&"_systemnote")
	          		End If
	          		%>
	          		</textarea>
					</td>
	  </tr>
	  <tr>
	    <td class="tdbg" height=23></td>
	    <td class="tdbg"><input type="submit" value="<%=strSubmit%>">(系统处于关闭状态时,以管理员身份仍能完整访问)</td>
	  </tr>
	</form>
</table>
<br>
<form name="systemcontro2" method="post" action="admin_index.asp?action=enmod">
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
		<td height=25 colspan=2 class="topbg"><strong>部分功能暂时关闭(服务器或者IIS重启后功能会自动开启)</strong></th>
	</tr>
  <tr>
  	<td class="tdbg" colsapn=2  height=23>如果取消禁止则需要去除勾选重新确定</td>
  </tr>
  <tr>
    <td class="tdbg" class="" colspan=2  height=23>
    	<%
    	Dim enStr0,enStr1,enStr2,enStr3,enStr4
    		enStr0=Application(cache_name_user&"_systemenmod")
    		If enStr0<>"" Then
    			enStr0=Split(enStr0,",")
    			If enStr0(0)="1" Then enStr1=" checked"
    			If enStr0(1)="1" Then enStr2=" checked"
    			If enStr0(2)="1" Then enStr3=" checked"
    			If enStr0(3)="1" Then enStr4=" checked"
    		End If
    	%>
    	<input type="checkbox" name="encomment" value="1" <%=enStr1%>>禁止回复与留言
<!--    	<input type="checkbox" name="enargue"  value="1"  <%=enStr2%>>禁止参与辩论-->
    	<input type="checkbox" name="enblog"  value="1" <%=enStr3%>>禁止发布日志
    	<input type="checkbox" name="entb"  value="1" <%=enStr4%>>禁止接收引用通告
    	<input type="submit" value="确定">
    	</td>
  </tr>
</table>
</form>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 class="topbg"><strong>oBlog 帮 助</strong>
  <tr>
    <td height=23 class="tdbg">1、<strong>将用户前台屏蔽以后此用户发布的所有文章(包括照片)都不会在首页被调用</strong>。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">2、<a href="admin_skin_help.asp" target="_blank"><font color="red">系统模板及用户模板的标记说明请点击这里。</font></a></td>
  </tr>
  <tr>
    <td class="tdbg">3、将用户锁定以后，此用户的blog页面也将被屏蔽。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">4、将IP屏蔽以后，此IP用户将不能登陆，且不能发表评论及留言。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">5、将博客设置为推荐，必须在后台修改用户资料才能实现。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">6、若上传文件不正常，请检查是否文件尺寸过大及服务器是否支持fso。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">7、有任何问题，请咨询oBlog官方网站<a href="http://www.oBlog.cn" target="_blank">http://www.oBlog.cn</a>。</td>
  </tr>
</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 colspan=2 class="topbg"><strong>oBlog开发</th></strong>
  <tr>
    <td width="20%" class="tdbg" height=23>程序开发：</td>
    <td width="80%" class="tdbg">  北京傲博致远计算机技术有限公司  威海研发中心 </td>
  </tr>
  <tr>
    <td class="tdbg" height=23>联系方式：</td>
    <td class="tdbg"> 威海研发中心网址：<a href="http://www.oblog.cn" target"=_blank">http://www.oblog.cn</a> 邮箱：webmaster@oblog.cn</td>
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
<script language=javascript src="http://www.oblog.cn/oblog4ver.asp"></script>
<%
end sub
%>