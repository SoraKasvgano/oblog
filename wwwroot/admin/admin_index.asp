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
		EventLog "��ʱ�ر��˱�վ",oblog.NowUrl&"?"&Request.QueryString
	Else
		Application(cache_name_user&"_systemstate")="run"
		EventLog "�����˱�վ",oblog.NowUrl&"?"&Request.QueryString
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
	EventLog "��ʱ�رգ����¿������˱�վ�Ĳ��ֹ���",oblog.NowUrl&"?"&Request.QueryString
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
<title>oBlog--��̨����</title>
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
  <p>���������汾���ͣ�������ϵͳҪ��IE5�����ϰ汾����ʹ�ñ�ϵͳ��</p>
  </body>
</noframes>
</html>
<%
end sub

sub admin_top()
%>
<html>
<head>
<title>oBlog��̨����ҳ��</title>
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

td {FONT-SIZE: 9pt;COLOR: #000000; FONT-FAMILY: "����"}
img {filter:Alpha(opacity:100); chroma(color=#FFFFFF)}
</style>
<base target="main">
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" height="100%" border=0 cellpadding=0 cellspacing=0>
  <tr valign=middle>
    <td width=10></td>
	<td width=50><a href="admin_adminmodifypwd.asp">�޸�����</a></td>
     <td align="left" width="500"><span id="ob4news"></span></td>
    <td width="50" align="left"><a href="../index.asp" target="_blank">վ����ҳ</a></td>
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
<title>oBlog--��̨����</title>
<link rel="stylesheet" href="images/admin/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body style="overflow-x:hidden;text-align:left;">
<div id="logo"></div>
<div id="TabPage1">
<!--TabPage1-->
	<div class="left_top"></div>
	<ul class="left_conten" id="oblog_0">
		<li><a href="admin_index.asp?action=main" target="main"><strong>������ҳ</strong></a>|<a href="admin_login.asp?action=logout"  target=_top><strong>�˳�</strong></a></li>
		<li><a>�û�����<%= session("adminname") %></a></li>
		<li><a href="../<%=SYSFOLDER_MANAGER%>/m_index.asp" target=_blank>�������ݹ���</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_1)">
		<li class="left_top_left left">ϵͳ����</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_1">
		<li><a href="admin_setup.asp" target="main" id="Tab101" onClick="javascript:switchTab('TabPage1','Tab101');">��վ��Ϣ����</a></li>
		<li><a href="admin_com.asp" target="main" id="Tab102" onClick="javascript:switchTab('TabPage1','Tab102');">�������������</a></li>
		<li><a href="admin_syslog.asp" target="main" id="Tab405" onClick="javascript:switchTab('TabPage1','Tab405');">ϵͳ������־����</a></li>
		<li><a href="admin_userclass.asp" target="main" id="Tab103" onClick="javascript:switchTab('TabPage1','Tab103');">ϵͳ���ͷ������</a></li>
		<li><a href="admin_logclass.asp" target="main" id="Tab104" onClick="javascript:switchTab('TabPage1','Tab104');">ϵͳ��־�������</a></li>
		<li><a href="admin_logclass.asp?t=1" target="main" id="Tab105" onClick="javascript:switchTab('TabPage1','Tab105');">ϵͳ���������</a></li>
		<li><a href="admin_logclass.asp?t=2" target="main" id="Tab107" onClick="javascript:switchTab('TabPage1','Tab107');">ϵͳȺ��������</a></li>
		<li><a href="../admin_edit.asp?do=3" target="main" id="Tab106a" onClick="javascript:switchTab('TabPage1','Tab106a');">�޸�ע������</a>|<a href="admin_note.asp?action=do3" target="main" id="Tab106b" onClick="javascript:switchTab('TabPage1','Tab106b');">�ı�</a></li>
		<li><a href="admin_js.asp" target="main" id="Tab108" onClick="javascript:switchTab('TabPage1','Tab108');">JS���ù���</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_2)">
		<li class="left_top_left left">��������</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_2">
		<li><a href="admin_score.asp" target="main" id="Tab201" onClick="javascript:switchTab('TabPage1','Tab201');">��վ�����ƶ�</a></li>
		<li><a href="admin_ask.asp" target="main" id="Tab209" onClick="javascript:switchTab('TabPage1','Tab209');">�Զ�����֤�������</a></li>
		<li><a href="../admin_edit.asp?do=1" target="main" id="Tab202a" onClick="javascript:switchTab('TabPage1','Tab202a');">�޸���������</a>|<a href="admin_note.asp?action=do1" target="main" id="Tab202b" onClick="javascript:switchTab('TabPage1','Tab202b');">�ı�</a></li>
		<li><a href="../admin_edit.asp?do=2" target="main" id="Tab203a" onClick="javascript:switchTab('TabPage1','Tab203a');">�޸���վ����</a>|<a href="admin_note.asp?action=do2" target="main" id="Tab203b" onClick="javascript:switchTab('TabPage1','Tab203b');">�ı�</a></li>
		<li><a href="../admin_edit.asp?do=4" target="main" id="Tab204a" onClick="javascript:switchTab('TabPage1','Tab204a');">�û���̨֪ͨ</a>|<a href="admin_note.asp?action=do4" target="main" id="Tab204b" onClick="javascript:switchTab('TabPage1','Tab204b');">�ı�</a></li>
		<li><a href="admin_lockip.asp" target="main" id="Tab205" onClick="javascript:switchTab('TabPage1','Tab205');">����IP����</a></li>
		<li><a href="admin_report.asp" target="main" id="Tab2055" onClick="javascript:switchTab('TabPage1','Tab2055');">��ӳ�������</a></li>
		<li><a href="admin_count.asp" target="main" id="Tab206" onClick="javascript:switchTab('TabPage1','Tab206');">����ϵͳ����</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_3)">
		<li class="left_top_left left">������</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_3">
		<li><a href="admin_ad.asp" target="main" id="Tab301" onClick="javascript:switchTab('TabPage1','Tab301');">�û�ҳ�������</a></li>
		<li><a href="admin_teamad.asp" target="main" id="Tab302" onClick="javascript:switchTab('TabPage1','Tab302');">Ⱥ��ҳ�������</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_4)">
		<li class="left_top_left left">����Ա��������</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_4">
		<li><a href="admin_admin.asp?Action=Add" target="main" id="Tab401" onClick="javascript:switchTab('TabPage1','Tab401');">����µĹ���Ա</a></li>
		<li><a href="admin_role.asp?action=add" target="main" id="Tab402" onClick="javascript:switchTab('TabPage1','Tab402');">���ݹ���Ա�ּ�</a></li>
		<li><a href="admin_role.asp" target="main" id="Tab403" onClick="javascript:switchTab('TabPage1','Tab403');">���ݹ���Ա�б�</a></li>
		<li><a href="admin_admin.asp" target="main" id="Tab404" onClick="javascript:switchTab('TabPage1','Tab404');">ȫ������Ա�б�</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_5)">
		<li class="left_top_left left">�û��ȼ�������</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_5">
		<li><a href="admin_user.asp" target="main" id="Tab506" onClick="javascript:switchTab('TabPage1','Tab506');">ȫ���û�����</a></li>
		<li><a href="admin_groups.asp?action=add" target="main" id="Tab501" onClick="javascript:switchTab('TabPage1','Tab501');">�����û��ȼ�</a></li>
		<li><a href="admin_groups.asp" target="main" id="Tab502" onClick="javascript:switchTab('TabPage1','Tab502');">�����û��ȼ�</a></li>
		<li><a href="admin_rename.asp" target="main" id="Tab503" onClick="javascript:switchTab('TabPage1','Tab503');">�û�����</a></li>
		<li><a href="admin_userdir.asp" target="main" id="Tab504" onClick="javascript:switchTab('TabPage1','Tab504');">�û�Ŀ¼����</a></li>
		<li><a href="admin_user.asp?Action=Update" target="main" id="Tab505" onClick="javascript:switchTab('TabPage1','Tab505');">�����û���̬ҳ</a></li>
	</ul>
	<div class="left_end"></div>
	<ul class="left_top" onClick="menu(oblog_6)">
		<li class="left_top_left left">ģ�����</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_6">
		<li><a href="admin_sysskin.asp?action=addskin" target="main" id="Tab601a" onClick="javascript:switchTab('TabPage1','Tab601a');">���ϵͳģ��</a>|<a href="admin_sysskin.asp?action=showskin" target="main" id="Tab601b" onClick="javascript:switchTab('TabPage1','Tab601b');">����</a></li>
		<li><a href="admin_skin.asp?action=insys1" target="main" id="Tab602a" onClick="javascript:switchTab('TabPage1','Tab602a');">ϵͳģ�嵼��</a>|<a href="admin_skin.asp?action=outsys" target="main" id="Tab602b" onClick="javascript:switchTab('TabPage1','Tab602b');">����</a></li>
		<li><a href="admin_userskin.asp?action=skinclass" target="main" id="Tab603" onClick="javascript:switchTab('TabPage1','Tab603');">�û�ģ��������</a></li>
		<li><a href="admin_userskin.asp?action=addskin" target="main" id="Tab604a" onClick="javascript:switchTab('TabPage1','Tab604a');">����û�ģ��</a>|<a href="admin_userskin.asp?action=showskin&ispass=1" target="main" id="Tab604b" onClick="javascript:switchTab('TabPage1','Tab604b');">����</a></li>
		<li><a href="admin_skin.asp?action=inuser1" target="main" id="Tab605a" onClick="javascript:switchTab('TabPage1','Tab605a');">�û�ģ�嵼��</a>|<a href="admin_skin.asp?action=outuser" target="main" id="Tab605b" onClick="javascript:switchTab('TabPage1','Tab605b');">����</a></li>
		<li><a href="admin_teamskin.asp?action=skinclass" target="main" id="Tab608" onClick="javascript:switchTab('TabPage1','Tab608');">Ⱥ��ģ��������</a></li>
		<li><a href="admin_teamskin.asp?action=addskin" target="main" id="Tab606a" onClick="javascript:switchTab('TabPage1','Tab606a');">���Ⱥ��ģ��</a>|<a href="admin_teamskin.asp?action=showskin&ispass=1" target="main" id="Tab606b" onClick="javascript:switchTab('TabPage1','Tab606b');">����</a></li>
		<li><a href="admin_skin.asp?action=inteam1" target="main" id="Tab607a" onClick="javascript:switchTab('TabPage1','Tab607a');">Ⱥ��ģ�嵼��</a>|<a href="admin_skin.asp?action=outteam" target="main" id="Tab607b" onClick="javascript:switchTab('TabPage1','Tab607b');">����</a></li>
	</ul>
	<div class="left_end"></div>
<%
'�����SQL Server���ݿ�������ʾ�ý�
If IS_SQLDATA=0 Then
%>
	<ul class="left_top" onClick="menu(oblog_7)">
		<li class="left_top_left left">���ݿ����</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_7">
		<li><a href="admin_database.asp?Action=Compact" target="main" id="Tab701" onClick="javascript:switchTab('TabPage1','Tab701');">ѹ�����ݿ�</a></li>
		<li><a href="admin_database.asp?Action=Backup" target="main" id="Tab702" onClick="javascript:switchTab('TabPage1','Tab702');">�������ݿ�</a></li>
		<li><a href="admin_database.asp?Action=Restore" target="main" id="Tab703" onClick="javascript:switchTab('TabPage1','Tab703');">�ָ����ݿ�</a></li>
		<li><a href="admin_database.asp?Action=SpaceSize" target="main" id="Tab704" onClick="javascript:switchTab('TabPage1','Tab704');">ϵͳ�ռ�ռ��</a></li>
	</ul>
	<div class="left_end"></div>
<%End If%>
	<ul class="left_top" onClick="menu(oblog_8)">
		<li class="left_top_left left">ϵͳ����</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_8">
		<li><a href="admin_analyze.asp?action=year" target="main" id="Tab801" onClick="javascript:switchTab('TabPage1','Tab801');">������ȷ���</a></li>
		<li><a href="admin_analyze.asp?action=month" target="main" id="Tab802" onClick="javascript:switchTab('TabPage1','Tab802');">�����¶ȷ���</a></li>
		<li><a href="admin_analyze.asp?action=hour" target="main" id="Tab803" onClick="javascript:switchTab('TabPage1','Tab803');">����ʱ�η���</a></li>
	</ul>
	<div class="left_end"></div>
<!--/TabPage1-->
</div>
<div id="cnt"></div>
	<ul class="left_top" onClick="menu(oblog_9)">
		<li class="left_top_left left">oBlog��Ϣ</li>
		<li class="left_top_right right"> </li>
	</ul>
	<ul class="left_conten" id="oblog_9">
		<li>��Ȩ����:<a>����������Զ������������޹�˾</a></li>
		<li>��վ��ַ:<a href="http://www.oBlog.cn" target="_blank">oBlog.cn</a></li>
		<li>����֧��:<a href="http://bbs.oBlog.cn" target="_blank">oBlog��̳</a></li>
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
<title>oBlog--��̨����</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">oBlog��̨������ҳ</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="100%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
<td height=25 colspan=2 class="topbg"><strong>ϵ ͳ ״ ̬</strong></th></tr>
     <tr>
    <td width="20%" class="tdbg" height=23>��ǰϵͳ�汾</td>
    <td width="80%" class="tdbg">&nbsp;&nbsp;<font color=blue><%=ver%></font></td>
  </tr>
     <tr>
    <td width="20%" class="tdbg" height=23>����ϵͳ�汾</td>
    <td width="80%" class="tdbg">&nbsp;&nbsp;<span id="ob4version"></span></td>
  </tr>

     <tr>
    <td width="20%" class="tdbg" height=23>��ǰϵͳ״̬</td>
    <td width="80%" class="tdbg">&nbsp;&nbsp;<font color=red>
    <%
    	Dim strSubmit
    	If Application(cache_name_user&"_systemstate")="stop" Then
    		Response.Write "�ر���"
    		strSubmit="��������"
    	Else
    		Response.Write "����������"
    		strSubmit="�ر�ϵͳ"
    	End If
    %></font>
    </td>
  </tr>
   <form name="systemcontrol" method="post" action="admin_index.asp?action=state">
	  <tr>
	    <td width="20%" class="tdbg" height=23>�ر�ʱ��ʾ��</td>
	    <td width="80%" class="tdbg">
	          	<textarea cols=60 rows=5 name="systemnote"><%
	          		If Application(cache_name_user&"_systemnote")="" Then
	          			Response.Write "ϵͳ����ά�����Ժ󿪷�"
	          		Else
	          			Response.Write Application(cache_name_user&"_systemnote")
	          		End If
	          		%>
	          		</textarea>
					</td>
	  </tr>
	  <tr>
	    <td class="tdbg" height=23></td>
	    <td class="tdbg"><input type="submit" value="<%=strSubmit%>">(ϵͳ���ڹر�״̬ʱ,�Թ���Ա���������������)</td>
	  </tr>
	</form>
</table>
<br>
<form name="systemcontro2" method="post" action="admin_index.asp?action=enmod">
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
		<td height=25 colspan=2 class="topbg"><strong>���ֹ�����ʱ�ر�(����������IIS�������ܻ��Զ�����)</strong></th>
	</tr>
  <tr>
  	<td class="tdbg" colsapn=2  height=23>���ȡ����ֹ����Ҫȥ����ѡ����ȷ��</td>
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
    	<input type="checkbox" name="encomment" value="1" <%=enStr1%>>��ֹ�ظ�������
<!--    	<input type="checkbox" name="enargue"  value="1"  <%=enStr2%>>��ֹ�������-->
    	<input type="checkbox" name="enblog"  value="1" <%=enStr3%>>��ֹ������־
    	<input type="checkbox" name="entb"  value="1" <%=enStr4%>>��ֹ��������ͨ��
    	<input type="submit" value="ȷ��">
    	</td>
  </tr>
</table>
</form>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 class="topbg"><strong>oBlog �� ��</strong>
  <tr>
    <td height=23 class="tdbg">1��<strong>���û�ǰ̨�����Ժ���û���������������(������Ƭ)����������ҳ������</strong>��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">2��<a href="admin_skin_help.asp" target="_blank"><font color="red">ϵͳģ�弰�û�ģ��ı��˵���������</font></a></td>
  </tr>
  <tr>
    <td class="tdbg">3�����û������Ժ󣬴��û���blogҳ��Ҳ�������Ρ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">4����IP�����Ժ󣬴�IP�û������ܵ�½���Ҳ��ܷ������ۼ����ԡ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">5������������Ϊ�Ƽ��������ں�̨�޸��û����ϲ���ʵ�֡�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">6�����ϴ��ļ��������������Ƿ��ļ��ߴ���󼰷������Ƿ�֧��fso��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">7�����κ����⣬����ѯoBlog�ٷ���վ<a href="http://www.oBlog.cn" target="_blank">http://www.oBlog.cn</a>��</td>
  </tr>
</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 colspan=2 class="topbg"><strong>oBlog����</th></strong>
  <tr>
    <td width="20%" class="tdbg" height=23>���򿪷���</td>
    <td width="80%" class="tdbg">  ����������Զ������������޹�˾  �����з����� </td>
  </tr>
  <tr>
    <td class="tdbg" height=23>��ϵ��ʽ��</td>
    <td class="tdbg"> �����з�������ַ��<a href="http://www.oblog.cn" target"=_blank">http://www.oblog.cn</a> ���䣺webmaster@oblog.cn</td>
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