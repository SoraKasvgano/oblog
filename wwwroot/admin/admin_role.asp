<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/inc_control.asp"-->
<script language="javascript">
function selectall(stype)
{
	var obj = document.getElementsByTagName("input");
	for (var i = 0;i<obj.length ;i++ ){
		var e = obj[i];
		if (e.type == 'radio'){
			if (e.value == stype){
				e.checked = true;
			}
		}
	}
}
</script>
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
		<li class="main_top_left left">内容管理员管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr >
      <td width="70" height="30"><strong>管理导航：</strong></td>
      <td height="30"><a href="admin_role.asp">内容管理员等级列表</a> | <a href="admin_role.asp?action=add">添加内容管理员等级</a> | <a href="admin_admin.asp">管理员账号管理</a>   | <a href="admin_admin.asp?Action=Add">新增管理员</a> </td>
    </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%
'不需要分页
Dim Action,sAction,Roleid
Dim sName,sUserReg,sUserName,sUserAdmin,sUserGroup,sGroups,sUserUpdate,sSecondDomain,sClasses,sClasses2
Dim sModScript,sModArgue,sModMeet,sModSpecial,sUserNews,sSiteNews,sBlogStar,sSkinsys,sSkinUser,sSkinQQ
Dim sWords,sIP,sSiteCount,sUserBlog,sUserRBlog,sUserCmt,sUserMsg,sUserTag,sUserUpfiles,sListUpfiles
Dim sGroupUser,sGroupBlog,sUserALL,sUserAdd,sUserAlbumCmt
Dim rs,sql
Roleid=Request.QueryString("Roleid")
If RoleId<>"" Then RoleId=Int(RoleId)
action=Trim(Request("action"))
select Case LCase(action)
	 Case "add","edit"
		call ShowForm(Roleid)
	 Case "save"
		call Save()
	Case "add","edit"
		Call ShowForm
	Case "del"
		conn.execute("Delete From oblog_roles Where roleid=" & Roleid)
		EventLog "进行了删除内容管理员组的操作，目标内容管理员组ID为："&OB_IIF(Roleid,"无"),oblog.NowUrl&"?"&Request.QueryString
		Response.Redirect "admin_role.asp"
	Case Else
		Call ShowList
	End select

Sub ShowForm(Roleid)
dim rst
If Roleid<>"" Then
	Set rst=oblog.execute("select * from oblog_roles Where roleid=" &Roleid)
	If Not rst.Eof Then
		sWords = rst("r_words")
		sIP = rst("r_IP")
		sSiteCount = rst("r_site_count")
		Roleid=rst("roleid")
		sName=rst("r_name")
		sSkinSys=rst("r_skin_sys")
		sSkinQQ=rst("r_skin_qq")
		sSkinUser=rst("r_skin_user")
		sGroups=Replace(ob_IIF(rst("r_groups"),"")," ","")
		sClasses=Replace(ob_IIF(rst("r_classes1"),"")," ","")
		sClasses2=Replace(ob_IIF(rst("r_classes2"),"")," ","")
		sUserReg=rst("r_user_reg")
		sUserName=rst("r_user_name")
		sUserAdmin=rst("r_user_admin")
		sUserGroup=rst("r_user_group")
		sUserUpdate=rst("r_user_update")
		sSecondDomain=rst("r_second_domain")
		sModScript=rst("r_mod_script")
		sModArgue=rst("r_mod_argue")
		sModMeet=rst("r_mod_meet")
		sModSpecial=rst("r_mod_special")
		sUserNews=rst("r_user_news")
		sSiteNews=rst("r_site_news")
		sBlogStar=rst("r_blogstar")
		sUserBlog=rst("r_user_blog")
		sUserRBlog=rst("r_user_rblog")
		sUserCmt=rst("r_user_cmt")
		sUserMsg=rst("r_user_msg")
		sUserTag=rst("r_user_tag")
		sUserUpfiles=rst("r_user_upfiles")
		sListUpfiles=rst("r_list_upfiles")
		sGroupUser=rst("r_group_user")
		sGroupBlog=rst("r_group_blog")
		sUserALL=rst("r_user_all")
		sUserAdd=rst("r_user_add")
		sUserAlbumCmt=rst("r_album_comment")
	End If
	Set rst=Nothing
End If

%>
<script language="javascript">
	function check1(){
		if(document.form1.r_name.value=""){
			alert("等级名称必须填写");
			document.form1.r_name.focus();
			return false;
			}
		}
	//return true;
</script>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">内容管理员登记功能定义</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="post" action="admin_role.asp?action=save&roleid=<%=roleid%>" id="form1" name="form1" onSubmit=""return check1()">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" colspan=2>
      	默认权限<a href="#h2" onClick="hookDiv('hh1','')"><img src="images/ico_help.gif" border=0></a>（此处与内容管理员页面的选项一一对应）
      	<div id="hh1" style="display:none" name="h1">
  	  如果启用日志审核功能，则所有的内容管理员均可进行日志审核<br/>
      如果启用了注册审核,则所有的内容管理员均可以审核注册<br/>
      所有内容管理员均可以增加关键字,增加黑名单IP<br/>
      所有内容管理员均可以设置精华文章/推荐文章
    </div>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >等级名称</td>
      <td width="409" height="25" >
		<% Call EchoInput("r_name",40,50,sName)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>常规管理</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >关键字管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_words","","",sWords)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >限制IP管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_IP","","",sIP)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >用户后台通知</td>
      <td height="25" ><% Call EchoRadio("r_site_news","","",sSiteNews)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td height="25" >发送站内短信</td>
      <td height="25" ><% Call EchoRadio("r_user_news","","",sUserNews)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >更新系统数据</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_site_count","","",sSiteCount)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>内容管理</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >日志管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_blog","","",sUserBlog)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >可管理的日志分类<br/>(不选择则允许管理任意分类))</td>
      <td>
	  <%
	  Dim sTmp,sChecked,rstTmp,i
	  sChecked=""
	  Set rstTmp=conn.Execute("select id,classname From oblog_logclass Where depth=0 And idType=0")
	  If rstTmp.Eof Then
			Response.Write "还没有定义任何日志分类"
      Else
		  Do While Not rstTmp.Eof
		  	If InStr(","&sClasses&",","," & rstTmp("id") & ",")>0 Then
		  		sChecked=" checked"
			  Else
			  	sChecked=" "
			  End If
			Response.Write "<input type=""checkbox"" name=""class1"" value="""& rstTmp("id") &"""" & sChecked &">" & rstTmp("classname") & "&nbsp;"
			i=i+1
			If i Mod 5 =0 Then Response.Write "<br/>"
			rstTmp.MoveNext
		  Loop
	  End If
	  rstTmp.Close
	  %></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >回收站管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_rblog","","",sUserRBlog)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >评论管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_cmt","","",sUserCmt)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >留言管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_msg","","",sUserMsg)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >TAG管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_tag","","",sUserTag)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >相册评论管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_album_comment","","",sUserAlbumCmt)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>群组管理</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >群组管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_group_user","","",sGroupUser)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >群组内容管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_group_blog","","",sGroupBlog)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>上传管理</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >上传管理用户清单</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_upfiles","","",sUserUpfiles)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >上传管理文件清单</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_list_upfiles","","",sListUpfiles)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>用户管理</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >全部用户管理</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_all","","",sUserALL)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >可管理的用户组<br/>(不选择则允许管理任意分类)</td>
      <td>
	  <%
	  sChecked=""
	  Set rstTmp=conn.Execute("select groupid,g_name From oblog_groups ")
	  If rstTmp.Eof Then
			Response.Write "还没有定义任何用户组"
      Else
		  Do While Not rstTmp.Eof
		  	If InStr(","&sGroups&",","," & rstTmp("groupid") & ",")>0 Then
		  		sChecked=" checked"
			Else
			 	sChecked=" "
			End If
			Response.Write "<input type=""checkbox"" name=""groupid"" value="""& rstTmp("groupid")&""""& sChecked &">" & rstTmp("g_name") & "&nbsp;"
			i=i+1
			If i Mod 6 =0 Then Response.Write "<br/>"
			rstTmp.MoveNext
		  Loop
	  End If
	  rstTmp.Close
	  %></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >博客之星管理</td>
      <td width="409" height="25" > <% Call EchoRadio("r_blogstar","","",sBlogStar)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >用户改名</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_name","","",sUserName)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >进入用户管理后台</td>
      <td width="409" height="25" > <% Call EchoRadio("r_user_admin","","",sUserAdmin)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > 增加新用户</td>
      <td width="409" height="25" ><% Call EchoRadio("r_user_add","","",sUserAdd)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > 是否可以修改用户组</td>
      <td width="409" height="25" ><% Call EchoRadio("r_user_group","","",sUserGroup)%> </td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>模版管理</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >管理系统模板</td>
      <td height="25" ><% Call EchoRadio("r_skin_sys","","",sSkinSys)%> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >管理用户模板</td>
      <td height="25" ><% Call EchoRadio("r_skin_user","","",sSkinSys)%> </td>
    </tr>
<!--    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否管理辩论(审核+总结)</td>
      <td height="25" ><% Call EchoRadio("r_mod_argue","","",sModArgue)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否管理活动(审核+总结)</td>
      <td height="25" ><% Call EchoRadio("r_mod_meet","","",sModMeet)%> </td>
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >是否允许用户启用脚本(需要进行审核,每次修改后仍然需要审核)</td>
      <td height="25" ><% Call EchoRadio("r_mod_script","","",sModScript)%> </td>
    </tr>-->
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >&nbsp;</td>
      <td height="25" ><input type="radio" name="selectradio" id="selectradio" value="0" onclick="selectall(0);" />全部选否&nbsp;
	<input type="radio" name="selectradio" id="selectradio" value="1"  onclick="selectall(1);"/>全部选是&nbsp; </td>
    </tr>
	<tr>
      <td colspan=2 align="center">
	  <input type="submit" value="保存" class="submit"> <input type="reset" value="取消"></td>
    </tr>
	</table>
	</form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%
Set rs=Nothing
End Sub

Sub Save
	Dim rst
	sName=Request.Form("r_name")
	sSkinSys=Request.Form("r_skin_sys")
	sSkinQQ=Request.Form("r_skin_qq")
	sSkinUser=Request.Form("r_skin_user")
	sGroups=Replace(Request.Form("groupid")," ","")
	sClasses=Replace(Request.Form("class1")," ","")
	sClasses2=Replace(Request.Form("class2")," ","")
	sUserReg=Request.Form("r_user_reg")
	sUserName=Request.Form("r_user_name")
	sUserAdmin=Request.Form("r_user_admin")
	sUserGroup=Request.Form("r_user_group")
	sUserUpdate=Request.Form("r_user_update")
	sSecondDomain=Request.Form("r_second_domain")
	sModScript=Request.Form("r_mod_script")
	sModArgue=Request.Form("r_mod_argue")
	sModMeet=Request.Form("r_mod_meet")
	sModSpecial=Request.Form("r_mod_special")
	sUserNews=Request.Form("r_user_news")
	sSiteNews=Request.Form("r_site_news")
	sBlogStar=Request.Form("r_blogstar")
	sWords = Request.Form("r_words")
	sIP = Request.Form("r_IP")
	sSiteCount = Request.Form("r_site_count")
	sUserBlog=Request.Form("r_user_blog")
	sUserRBlog=Request.Form("r_user_rblog")
	sUserCmt=Request.Form("r_user_cmt")
	sUserMsg=Request.Form("r_user_msg")
	sUserTag=Request.Form("r_user_tag")
	sUserUpfiles=Request.Form("r_user_upfiles")
	sListUpfiles=Request.Form("r_list_upfiles")
	sGroupUser=Request.Form("r_group_user")
	sGroupBlog=Request.Form("r_group_blog")
	sUserALL=Request.Form("r_user_all")
	sUserAdd=Request.Form("r_user_add")
	sUserAlbumCmt=Request.Form("r_album_comment")
	If sName="" Then
		%>
		<script language="javascript">
			alert("等级名称必须填写");
			history.back();
		</script>
		<%
		Response.End
	End If
	Set rst=Server.CreateObject("Adodb.Recordset")
	If Roleid<>"" Then
		rst.Open "select * From oblog_roles Where roleid=" & Int(Roleid),conn,1,3
	Else
		rst.Open "select * From oblog_roles Where r_name='" & sName & "'",conn,1,1
		If Not rst.Eof Then
			rst.Close
			Set rst=Nothing
			%>
			<script language="javascript">
			alert("等级名称 "&sName&" 已经存在");
			history.back();
		</script>
			<%
			Response.End
		End If
		rst.Close
		rst.Open "select * From oblog_roles Where 1=0",conn,1,3
		'进行数据校验
		rst.AddNew
	End If
	rst("r_name")=sName
	rst("r_skin_sys")=sSkinSys
	rst("r_skin_qq")=sSkinQQ
	rst("r_skin_user")=sSkinUser
	rst("r_groups")=sGroups
	rst("r_classes1")=sClasses
	rst("r_classes2")=sClasses2
	rst("r_user_reg")=sUserReg
	rst("r_user_name")=sUserName
	rst("r_user_admin")=sUserAdmin
	rst("r_user_group")=sUserGroup
	rst("r_user_update")=sUserUpdate
	rst("r_second_domain")=sSecondDomain
	rst("r_mod_script")=sModScript
	rst("r_mod_argue")=sModArgue
	rst("r_mod_meet")=sModMeet
	rst("r_mod_special")=sModSpecial
	rst("r_user_news")=sUserNews
	rst("r_site_news")=sSiteNews
	rst("r_blogstar")=sBlogStar
	rst("r_words") =sWords
	rst("r_IP") =sIP
	rst("r_site_count") = sSiteCount
	rst("r_user_blog") = sUserBlog
	rst("r_user_rblog") = sUserRBlog
	rst("r_user_cmt") = sUserCmt
	rst("r_user_msg") = sUserMsg
	rst("r_user_tag") = sUserTag
	rst("r_user_upfiles")=sUserUpfiles
	rst("r_list_upfiles")=sListUpfiles
	rst("r_group_user")=sGroupUser
	rst("r_group_blog")=sGroupBlog
	rst("r_user_all")=sUserALL
	rst("r_user_add")=sUserAdd
	rst("r_album_comment")=sUserAlbumCmt
	rst.Update
	Set rst=Nothing
	EventLog "进行了添加(修改)内容管理员组权限的操作，目标内容管理员组ID为："&OB_IIF(Roleid,"无"),oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "保存成功",""
End Sub

Sub ShowList()
	Dim rstM,rst
	Set rst=conn.Execute("select roleid,r_name From oblog_roles Order By roleid")
	Set rstM=conn.Execute("select * From oblog_admin Where roleid>0")
	%>
<br />
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">内容管理员等级列表</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
	<%
	If rst.Eof Then
	%>
		<tr><td align="center">目前还没有定义任何管理员等级信息</td></tr>
	<%
		rst.Close
		Set rst=Nothing
		Exit Sub
	End If
	%>
	<tr><td>等级编号</td><td>等级名称<td>账号列表</td><td>操作</td>
	<%
	Do While Not rst.Eof
		rstM.Filter="roleid=" & rst(0)
	%>
		<tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
		<td><%=rst(0)%></td>
		<td><%=rst(1)%></td>
		<td><%
			'写列表
			If rstM.Eof Then
				Response.Write "还没有分配"
			Else
				Do While Not rstM.Eof
					If rstM("userid")<>"" Then
						Response.Write rstM("username")& "("&rstM("userid")&")<br/>"
					Else
						Response.Write rstM("username") & "<br/>"
					End If
					rstM.Movenext
				Loop
			End If
			%></td>
		<td><a href="admin_role.asp?action=edit&roleid=<%=rst(0)%>">修改</a>&nbsp;&nbsp;
			<a href="admin_role.asp?action=del&roleid=<%=rst(0)%>" onClick="javascript:if(confirm('确认要删除该等级吗?')==false)return false;">删除</a>
		</td>
		</tr>
	<%
		rst.MoveNext
	Loop
	rst.Close
	rstM.Close
	Set rst=Nothing
	Set rstM=Nothing
	%>
	</tr></table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
	<%
End Sub
Set oblog = Nothing
%>
</body>
</html>