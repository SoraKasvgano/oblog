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
<title>oBlog--��̨����</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">���ݹ���Ա����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr >
      <td width="70" height="30"><strong>��������</strong></td>
      <td height="30"><a href="admin_role.asp">���ݹ���Ա�ȼ��б�</a> | <a href="admin_role.asp?action=add">������ݹ���Ա�ȼ�</a> | <a href="admin_admin.asp">����Ա�˺Ź���</a>   | <a href="admin_admin.asp?Action=Add">��������Ա</a> </td>
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
'����Ҫ��ҳ
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
		EventLog "������ɾ�����ݹ���Ա��Ĳ�����Ŀ�����ݹ���Ա��IDΪ��"&OB_IIF(Roleid,"��"),oblog.NowUrl&"?"&Request.QueryString
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
			alert("�ȼ����Ʊ�����д");
			document.form1.r_name.focus();
			return false;
			}
		}
	//return true;
</script>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">���ݹ���Ա�Ǽǹ��ܶ���</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="post" action="admin_role.asp?action=save&roleid=<%=roleid%>" id="form1" name="form1" onSubmit=""return check1()">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" colspan=2>
      	Ĭ��Ȩ��<a href="#h2" onClick="hookDiv('hh1','')"><img src="images/ico_help.gif" border=0></a>���˴������ݹ���Աҳ���ѡ��һһ��Ӧ��
      	<div id="hh1" style="display:none" name="h1">
  	  ���������־��˹��ܣ������е����ݹ���Ա���ɽ�����־���<br/>
      ���������ע�����,�����е����ݹ���Ա���������ע��<br/>
      �������ݹ���Ա���������ӹؼ���,���Ӻ�����IP<br/>
      �������ݹ���Ա���������þ�������/�Ƽ�����
    </div>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�ȼ�����</td>
      <td width="409" height="25" >
		<% Call EchoInput("r_name",40,50,sName)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>�������</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�ؼ��ֹ���</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_words","","",sWords)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >����IP����</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_IP","","",sIP)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�û���̨֪ͨ</td>
      <td height="25" ><% Call EchoRadio("r_site_news","","",sSiteNews)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td height="25" >����վ�ڶ���</td>
      <td height="25" ><% Call EchoRadio("r_user_news","","",sUserNews)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >����ϵͳ����</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_site_count","","",sSiteCount)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>���ݹ���</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��־����</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_blog","","",sUserBlog)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�ɹ������־����<br/>(��ѡ������������������))</td>
      <td>
	  <%
	  Dim sTmp,sChecked,rstTmp,i
	  sChecked=""
	  Set rstTmp=conn.Execute("select id,classname From oblog_logclass Where depth=0 And idType=0")
	  If rstTmp.Eof Then
			Response.Write "��û�ж����κ���־����"
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
      <td width="348" height="25" >����վ����</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_rblog","","",sUserRBlog)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >���۹���</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_cmt","","",sUserCmt)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >���Թ���</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_msg","","",sUserMsg)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >TAG����</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_tag","","",sUserTag)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >������۹���</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_album_comment","","",sUserAlbumCmt)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>Ⱥ�����</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >Ⱥ�����</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_group_user","","",sGroupUser)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >Ⱥ�����ݹ���</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_group_blog","","",sGroupBlog)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>�ϴ�����</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�ϴ������û��嵥</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_upfiles","","",sUserUpfiles)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�ϴ������ļ��嵥</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_list_upfiles","","",sListUpfiles)%></td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>�û�����</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >ȫ���û�����</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_all","","",sUserALL)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�ɹ�����û���<br/>(��ѡ������������������)</td>
      <td>
	  <%
	  sChecked=""
	  Set rstTmp=conn.Execute("select groupid,g_name From oblog_groups ")
	  If rstTmp.Eof Then
			Response.Write "��û�ж����κ��û���"
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
      <td width="348" height="25" >����֮�ǹ���</td>
      <td width="409" height="25" > <% Call EchoRadio("r_blogstar","","",sBlogStar)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�û�����</td>
      <td width="409" height="25" >
		<% Call EchoRadio("r_user_name","","",sUserName)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�����û������̨</td>
      <td width="409" height="25" > <% Call EchoRadio("r_user_admin","","",sUserAdmin)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > �������û�</td>
      <td width="409" height="25" ><% Call EchoRadio("r_user_add","","",sUserAdd)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > �Ƿ�����޸��û���</td>
      <td width="409" height="25" ><% Call EchoRadio("r_user_group","","",sUserGroup)%> </td>
    </tr>
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>ģ�����</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >����ϵͳģ��</td>
      <td height="25" ><% Call EchoRadio("r_skin_sys","","",sSkinSys)%> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�����û�ģ��</td>
      <td height="25" ><% Call EchoRadio("r_skin_user","","",sSkinSys)%> </td>
    </tr>
<!--    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ�������(���+�ܽ�)</td>
      <td height="25" ><% Call EchoRadio("r_mod_argue","","",sModArgue)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ����(���+�ܽ�)</td>
      <td height="25" ><% Call EchoRadio("r_mod_meet","","",sModMeet)%> </td>
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�Ƿ������û����ýű�(��Ҫ�������,ÿ���޸ĺ���Ȼ��Ҫ���)</td>
      <td height="25" ><% Call EchoRadio("r_mod_script","","",sModScript)%> </td>
    </tr>-->
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >&nbsp;</td>
      <td height="25" ><input type="radio" name="selectradio" id="selectradio" value="0" onclick="selectall(0);" />ȫ��ѡ��&nbsp;
	<input type="radio" name="selectradio" id="selectradio" value="1"  onclick="selectall(1);"/>ȫ��ѡ��&nbsp; </td>
    </tr>
	<tr>
      <td colspan=2 align="center">
	  <input type="submit" value="����" class="submit"> <input type="reset" value="ȡ��"></td>
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
			alert("�ȼ����Ʊ�����д");
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
			alert("�ȼ����� "&sName&" �Ѿ�����");
			history.back();
		</script>
			<%
			Response.End
		End If
		rst.Close
		rst.Open "select * From oblog_roles Where 1=0",conn,1,3
		'��������У��
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
	EventLog "���������(�޸�)���ݹ���Ա��Ȩ�޵Ĳ�����Ŀ�����ݹ���Ա��IDΪ��"&OB_IIF(Roleid,"��"),oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "����ɹ�",""
End Sub

Sub ShowList()
	Dim rstM,rst
	Set rst=conn.Execute("select roleid,r_name From oblog_roles Order By roleid")
	Set rstM=conn.Execute("select * From oblog_admin Where roleid>0")
	%>
<br />
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">���ݹ���Ա�ȼ��б�</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
	<%
	If rst.Eof Then
	%>
		<tr><td align="center">Ŀǰ��û�ж����κι���Ա�ȼ���Ϣ</td></tr>
	<%
		rst.Close
		Set rst=Nothing
		Exit Sub
	End If
	%>
	<tr><td>�ȼ����</td><td>�ȼ�����<td>�˺��б�</td><td>����</td>
	<%
	Do While Not rst.Eof
		rstM.Filter="roleid=" & rst(0)
	%>
		<tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
		<td><%=rst(0)%></td>
		<td><%=rst(1)%></td>
		<td><%
			'д�б�
			If rstM.Eof Then
				Response.Write "��û�з���"
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
		<td><a href="admin_role.asp?action=edit&roleid=<%=rst(0)%>">�޸�</a>&nbsp;&nbsp;
			<a href="admin_role.asp?action=del&roleid=<%=rst(0)%>" onClick="javascript:if(confirm('ȷ��Ҫɾ���õȼ���?')==false)return false;">ɾ��</a>
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