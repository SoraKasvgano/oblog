<!--#include file="inc/inc_sys.asp"-->
<%
Dim sAction,rst,ErrMsg
Dim sGid,sName,sLevel,sStyle,sPoints,sAutoUpdate,sOBCodes,sDayPosts,sPostChk,sSecondDomain,sDomain,sClasses,sSkins,sTeamSkins
Dim sSkinEdit,sSkinScript,sPostScript,sADSystem,sADuser,sQQNumber,sModNote,sModAddress,sModZhai,sModArgue,sModActions,sModScript,sIn_group
Dim sUpTypes,sUpSize,sUpSpace,sUpWaterMark,sIsPassword,sPMs,sUpdates,sCodePost,sMailPost,sDownLoad,sGetPwd,sDigg
Link_Database
'Layout
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�û��ȼ�����</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�û��ȼ�����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="tdbg">
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30">
    	<a href="admin_groups.asp">��ǰ���б�</a>&nbsp;|&nbsp;
    	<a href="admin_groups.asp?action=add">�����û��ȼ�</a>&nbsp;|&nbsp;
    	<a href="admin_groups.asp?action=move">��ͬ�ȼ����û�ת��</a>
    </td>
  </tr>
  <tr class="tdbg">
  <td height="30" colspan="2" >
    	<font color="red"><strong>�ȼ�������С���鱻Ĭ��Ϊ�û�ע���������Ĭ�ϵȼ�</strong></font>
   </td>
  </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<br />
<%
sGid=Request.QueryString("gid")
sGid=CheckInt(sGid,"")
sAction=LCase(Request.QueryString("action"))
select Case sAction
	Case "add","show"
		Call ShowForm(sGid)
	Case "move"
		Call ShowMove()
	Case "savemove"
  		Call SaveMove
 	Case "save"
		Call Save
	Case "del"
		Set rst=conn.execute("select Count(Userid) From oblog_user Where user_group=" & sGid)
		If rst.Eof Then
			conn.execute("Delete From oblog_groups Where groupid=" & sGid)
		Else
			If rst(0)>0 Then
				ErrMsg="�õȼ�����"&rst(0)&"�û�����,���Ƚ��ò����û�ת�Ƶ������ȼ�,�˴��г����õȼ���ĵȼ��б�,����ѡ��"

			Else
				conn.execute("Delete From oblog_groups Where groupid=" & sGid)
				ErrMsg="�õȼ��Ѿ���ɾ��"
			End If
		End If
		EventLog "�����û����ɾ��������Ŀ���û���IDΪ"&sGid,oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "ɾ���ɹ�","admin_groups.asp"
	Case "restat"
		Set rst = oblog.execute ("select count(userid) from oblog_user where user_group = "&sGid)
		oblog.execute ("update oblog_groups set g_members ="&rst(0)&" where groupid = " & sGid)
		rst.close
		Set rst=Nothing
		oblog.ShowMsg "����ͳ�Ƴɹ�","admin_groups.asp"
	Case  Else
		Call ShowList
End select

conn.Close
Set conn=Nothing
Dim sJs
sJs=JsValid("form1","g_name","1",1,50,"��������ȼ�����!")
sJs=sJs & JsValid("form1","g_uptypes","1",1,100,"������д�ϴ��ļ����ͣ�")
Call MakeValidJs("form1","checksubmit",sJs)
%>
</body>
</html>


<%
'Biz
Function GetMaxOrder()
	 Dim rst
	 Set rst=conn.execute("select Max(g_order) From oblog_groups")
	 If rst.Eof Or IsNull(rst(0)) Or Not IsNumeric(rst(0)) Then
	 	GetMaxOrder =1
	 Else
	  	GetMaxOrder =Int(rst(0))+1
	 End If
End Function

Sub ShowList()
	Set rst=conn.Execute("select groupid,g_name,g_points,g_autoupdate,g_members,g_level From oblog_groups Order By g_level")
	%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�û��ȼ���Ϣ</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
	<%
	If rst.Eof Then
	%>
		<tr><td align="center">Ŀǰ��û�ж����κ��û��ȼ���Ϣ</td></tr>
	<%
		rst.Close
		Set rst=Nothing
		Exit Sub
	End If
	%>
	<tr><td>�ȼ�����</td><td>�ȼ�����<td>��������</td><td>�Ƿ��Զ�����</td><td>�û���</td><td>����</td>
	<%
	Do While Not rst.Eof
	%>
		<tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
		<td><%=rst(5)%></td>
		<td><%=rst(1)%></td>
		<td><%=rst(2)%></td>
		<td><%If rst(3)="1" Then Response.Write "��" Else Response.Write "��"%></td>
		<td><%=rst(4)%></td>
		<td><a href="admin_groups.asp?action=restat&gid=<%=rst(0)%>">����ͳ��</a>&nbsp;&nbsp;<a href="admin_groups.asp?action=show&gid=<%=rst(0)%>">�޸�</a>&nbsp;&nbsp;
		<%If rst(4)=0 Then%>
			<a href="admin_groups.asp?action=del&gid=<%=rst(0)%>" onclick="return confirm('ȷ��Ҫɾ��ѡ�е��û�����');">ɾ��</a>
		<%End If%>
		</td>
		</tr>
	<%
		rst.MoveNext
	Loop
	rst.Close
	Set rst=Nothing
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

Sub ShowForm(Gid)
Dim rst,iOrder

If Gid<>"" And Gid>0 Then
	Set rst=oblog.execute("select * From oblog_groups Where groupid=" &Gid)
	If Not rst.Eof Then
		sGid=rst("groupid")
		sName=rst("g_name")
		sLevel=rst("g_level")
		sStyle=rst("g_style")
		sPoints=CheckInt(rst("g_points"),"0")
		sAutoUpdate=rst("g_autoupdate")
		sOBCodes=rst("g_obcodes")
		sDayPosts=rst("g_post_day")
		sPostChk=rst("g_post_chk")
		sSecondDomain=rst("g_seconddomain")
		sDomain=rst("g_domain")
		sClasses=rst("g_classes")
		sSkins=rst("g_skins")
'		sTeamSkins=rst("g_team_skins")
		sTeamSkins=rst("g_style")
		sSkinEdit=rst("g_skin_edit")
		sSkinScript=rst("g_skin_script")
		sPostScript=rst("g_post_script")
		sADSystem=rst("g_ad_sys")
		sADuser=rst("g_ad_user")
		sQQNumber=rst("g_qq_number")
		sModNote=rst("g_mod_note")
		sModAddress=rst("g_mod_address")
		sModZhai=rst("g_mod_zhai")
		sModArgue=rst("g_mod_argue")
		sModActions=rst("g_mod_meet")
		sUpTypes=rst("g_up_types")
		sUpSize=rst("g_up_size")
		sUpSpace=rst("g_up_space")
		sUpWaterMark=rst("g_up_watermark")
		sIsPassword=rst("g_is_password")
		sPMs=rst("g_pm_numbers")
		sCodePost=OB_IIF(rst("is_code_addblog"),"0")
		sUpdates=rst("oneday_update")
		sIn_group=rst("in_group")
		If oblog.cacheconfig(51)="1" Then
			sMailPost=OB_IIF(rst("g_mailpost"),"0")
		End If
		sDownLoad = rst("g_download")
		sGetPwd = rst("g_getpwd")
	End If
	If Trim(sPoints)="" Then sPoints="0"
	rst.Close
	Set rst=Nothing
Else
	Set rst=oblog.execute("select Max(g_level) From oblog_groups")
	If Not rst.Eof Then
		sLevel=OB_IIF(rst(0),0)
	Else
		sLevel=0
	End If
	sLevel=sLevel+1
End If

%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�ȼ���Ϣ����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="post" action="admin_groups.asp?action=save&gid=<%=Gid%>" id="form1" name="form1" onSubmit="return checksubmit();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
     <td width="348" height="25" >�ȼ�����</td>
     <td height="25" ><% Call EchoInput("g_name",40,40,sName)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�ȼ�����(�������ظ�)��</td>
      <td height="25" ><% Call EchoInput("g_level",40,40,sLevel)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�������õȼ�����С���֣�</td>
      <td height="25" ><% Call EchoInput("g_points",40,10,sPoints)%>(����õȼ��������������뽫������Ϊ0)</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > �Ƿ������Զ�����</td>
      <td> <% Call EchoRadio("g_autoupdate","","",sAutoupdate)%>  </td>
	  </tr>
	  <!--
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >���css��ʽ,ע���ڳ����н�ʹ�� style="xxx" ����ʽ</td>
      <td height="25" ><% Call EchoInput("g_font",40,20,sStyle)%></td>
    </tr>
    -->
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ������޸Ķ�������</td>
      <td height="25" ><% Call EchoRadio("g_seconddomain","","",sSecondDomain)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ���԰󶨶�������</td>
      <td height="25" ><% Call EchoRadio("g_domain","","",sDomain)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >ÿ�����ɷ�����������Ŀ��0Ϊ��������</td>
      <td height="25" ><% Call EchoInput("g_obcodes",40,40,sOBCodes)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��������־�Ƿ���Ҫ���</td>
      <td height="25" > <% Call EchoRadio("g_logchk","","",sPostChk)%></td>
    </tr>
	 <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >��������һ����־����<br/>(��ѡ�������������������)</td>
      <td>
	  <%
	  Dim sTmp,sChecked,rstTmp,i
	  sTmp="," & sClasses & ","
	  sChecked=""
	  Set rstTmp=conn.Execute("select * From oblog_logclass Where depth=0 And idType=0")
	  If rstTmp.Eof Then
			Response.Write "��û�ж����κη���"
      Else
		  Do While Not rstTmp.Eof
		  	If InStr(sTmp,"," & rstTmp("id") & ",")>0 Then
		  		sChecked=" checked"
		  	Else
		  		sChecked=""
		  	End If
			Response.Write "<input name=""g_classes"" type=""checkbox"" value="& rstTmp("id")& sChecked &">" & rstTmp("classname") & "&nbsp;"
			i=i+1
			If i Mod 5 =0 Then Response.Write "<br/>"
			rstTmp.MoveNext
		  Loop
	  End If
	  rstTmp.Close
	  %></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >ÿ������������־��Ŀ</td>
      <td><% Call EchoInput("g_dayposts",40,40,sDayPosts)%>(0Ϊ������)</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >��־���Ƿ�����ʹ�ýű�</td>
      <td height="25" > <% Call EchoRadio("g_post_script","","",sPostScript)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >����ʹ�õ�Ⱥ��ģ�����<br/>(���Ϊ��ѡ��������ʹ������ģ��,��ͬ)</td>
      <td height="25" >
	  <%
	  sTmp=Replace("," & sTeamSkins & ","," ","")
	  sChecked=""
	  i=0
	  Set rstTmp=conn.Execute("select * From oblog_skinclass Where iType=1")
	  If rstTmp.Eof Then
			Response.Write "��û�ж����κη���"
      Else

		  Do While Not rstTmp.Eof
		  	If InStr(sTmp,"," & rstTmp("classId") & ",")>0 Then
		  		sChecked=" checked"
		  	Else
		  		sChecked=""
		  	End If
			Response.Write "<input name=""g_team_skins"" type=""checkbox"" value="& rstTmp("classId")& sChecked &">" & rstTmp("classname") & "("&rstTmp("iCount")&")&nbsp;"
			i=i+1
			If i Mod 6 =0 Then Response.Write "<br/>"
			rstTmp.MoveNext
		  Loop
	  End If
	  rstTmp.Close
	  Set rstTmp=Nothing
	  %>
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >����ʹ�õ��û�ģ�����</td>
      <td height="25" >
	  <%
	  sTmp=Replace("," & sSkins & ","," ","")
	  sChecked=""
	  i=0
	  Set rstTmp=conn.Execute("select * From oblog_skinclass Where iType=0")
	  If rstTmp.Eof Then
			Response.Write "��û�ж����κη���"
      Else

		  Do While Not rstTmp.Eof
		  	If InStr(sTmp,"," & rstTmp("classId") & ",")>0 Then
		  		sChecked=" checked"
		  	Else
		  		sChecked=""
		  	End If
			Response.Write "<input name=""g_skins"" type=""checkbox"" value="& rstTmp("classId")& sChecked &">" & rstTmp("classname") & "("&rstTmp("iCount")&")&nbsp;"
			i=i+1
			If i Mod 6 =0 Then Response.Write "<br/>"
			rstTmp.MoveNext
		  Loop
	  End If
	  rstTmp.Close
	  Set rstTmp=Nothing
	  %>
	  </td>
    </tr>
	  <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ�����༭ģ��</td>
      <td height="25" > <% Call EchoRadio("g_skin_edit","","",sSkinEdit)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >ģ�����Ƿ�ʹ�ýű�</td>
      <td height="25" > <% Call EchoRadio("g_skin_script","","",sSkinScript)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��������Ⱥ����Ŀ(0Ϊ��������)</td>
     <td><% Call EchoInput("g_qqnumbers",40,40,sQQNumber)%></td>
    </tr>
    <!--
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ����ʹ��ͨѶ¼</td>
       <td height="25" > <% Call EchoRadio("g_mod_address","","",sModAddress)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ�����ʹ����ժ</td>
       <td height="25" > <% Call EchoRadio("g_mod_Zhai","","",sModZhai)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ�������˼��±�</td>
       <td height="25" > <% Call EchoRadio("g_mod_note","","",sModNote)%></td>
    </tr>
    -->
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�ϴ��ռ�Ĵ�С(0Ϊ������,-1Ϊ������)</td>
      <td height="25" ><% Call EchoInput("g_UpSpace",40,40,sUpSpace)%>KB</td>
    </tr>
	 <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��һ�ļ�����</td>
      <td height="25" ><% Call EchoInput("g_UpSize",40,40,sUpSize)%>KB</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�ϴ��ļ�����</td>
      <td height="25" ><% Call EchoInput("g_uptypes",50,100,sUpTypes)%>(��|�ָ�)</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >ͼƬ�Ƿ��ˮӡ</td>
      <td height="25" > <% Call EchoRadio("g_upwatermark","","",sUpWaterMark)%> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ�������վ����</td>
      <td height="25" > <% Call EchoRadio("g_is_password","","",sIsPassword)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�û�����Ϣ��Ŀ����</td>
      <td height="25" ><% Call EchoInput("g_pms",5,5,sPMs)%></td>
    </tr>

    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ�ǿ����ʾ���</td>
      <td height="25" ><% Call EchoRadio("g_ad_system","","",sAdSystem)%></td>
    </tr>
	    <!--
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ������û������Լ��Ĺ��</td>
      <td height="25" > <% Call EchoRadio("g_ad_user","","",sAdUser)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ������û����ýű�(��Ҫ�������,ÿ���޸ĺ���Ȼ��Ҫ���)</td>
      <td height="25" > <% Call EchoRadio("g_modscript","","",sModScript)%></td>
    </tr>
      <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ������û������������</td>
      <td height="25" > <% Call EchoRadio("g_modArgue","","",sModArgue)%></td>
    </tr>

	 <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ������û���������</td>
      <td height="25" > <% Call EchoRadio("g_modactions","","",sModActions)%></td>
    </tr>
    -->
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td height="25" >�û�������־�Ƿ���Ҫ��֤�룺</td>
      <td height="25" ><% Call EchoRadio("is_code_addblog","","",sCodePost)%></td>
    </tr>
     <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >ÿ�������û�����ȫվ���ٴΣ�</td>
      <td height="25" ><% Call EchoInput("oneday_update",5,5,sUpdates)%>��(����Ϊ0�򲻽�������)</td>
    </tr>
	  <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >��������Ⱥ������</td>
      <td height="25" ><% Call EchoInput("sIn_group",5,5,sIn_group)%>��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ��������ظ�����</td>
      <td height="25" ><% Call EchoRadio("g_download","","",sDownLoad)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ�����ȡ�����룺</td>
      <td height="25" ><% Call EchoRadio("g_getpwd","","",sGetPwd)%></td>
    </tr>
    <%If oblog.CacheConfig(51)="1" Then%>
     <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ�����ʹ���ʼ�����ŷ�����־��</td>
      <td height="25" ><% Call EchoRadio("g_mailpost","","",sMailPost)%></td>
    </tr>
    <%End If%>
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
End Sub

Sub Save()
	On Error Resume Next
	Dim Sql,tsql

	sGid=CheckInt(Request("gid"),"")
	sName=CheckStr(Request.Form("g_name"),20,"0")
	sLevel=CheckInt(Request.Form("g_level"),"1")
	sStyle=CheckStr(Request.Form("g_style"),"10","0")
	sPoints=CheckInt(Request.Form("g_points"),"")
	sAutoUpdate=Check01(Request.Form("g_autodate"),"1")
	sOBCodes=CheckInt(Request.Form("g_obcodes"),"2")
	sDayPosts=CheckInt(Request.Form("g_dayposts"),"10")
	sPostChk=Check01(Request.Form("g_logchk"),0)
	sSecondDomain=Check01(Request.Form("g_seconddomain"),1)
	sAutoUpdate=Check01(Request.Form("g_autoupdate"),1)
	sDomain=Check01(Request.Form("g_domain"),0)
	sClasses=FilterIds(Request.Form("g_classes"))
	sSkins=Request.Form("g_skins")
	sTeamSkins=Request.Form("g_team_skins")
	sSkinEdit=Check01(Request.Form("g_skin_edit"),0)
	sSkinScript=Check01(Request.Form("g_skin_script"),0)
	sPostScript=Check01(Request.Form("g_post_script"),0)
	sADSystem=Check01(Request.Form("g_ad_system"),0)
	sADuser=Check01(Request.Form("g_ad_user"),0)
	sQQNumber=CheckInt(Request.Form("g_qqnumbers"),3)
	sModNote=Check01(Request.Form("g_mod_note"),1)
	sModAddress=Check01(Request.Form("g_mod_address"),0)
	sModZhai=Check01(Request.Form("g_mod_zhai"),1)
	sModArgue=Check01(Request.Form("g_mod_aruge"),1)
	sModActions=Check01(Request.Form("g_mod_action"),0)
	sUpTypes=Request.Form("g_uptypes")
	sUpSize=CheckInt(Request.Form("g_upsize"),"")
	sUpSpace=CheckInt(Request.Form("g_upspace"),"")
	sUpWaterMark=Check01(Request.Form("g_upwatermark"),0)
	sIsPassword=Check01(Request.Form("g_is_password"),0)
	sPMs=Int(OB_IIF(Request.Form("g_pms"),30))
	sUpdates=OB_IIF(Request.Form("oneday_update"),5)
	sIn_group=OB_IIF(Request.Form("sIn_group"),10)
	sCodePost=Check01(Request.Form("is_code_addblog"),0)
	sMailPost=Check01(Request.Form("g_mailpost"),0)
	sDownLoad = Check01(Request.Form("g_download"),0)
	sGetPwd = Check01(Request.Form("g_getpwd"),1)
	'��������У��
	If sName="" Then oblog.AddErrStr("������д������")
  	If sLevel="" Or Not IsNumeric(sLevel)   Then oblog.AddErrStr("<li>������д�鼶��</li>")
	If sPoints="" Then  sPoints="0"
	If oblog.ErrStr<>"" Then Response.Write oblog.ErrStr:Response.End
	sql="select * From oblog_groups"
	If sGid>0 Then
		sql=sql & " Where groupid=" & sGid
		tsql=" AND groupid<>" & sGid
	Else
		sql=sql & " Where 1=0"
	End If
	'Check Group Name
	Set rst=conn.Execute("select groupid From oblog_groups Where g_name='" & sName & "'"& tsql)
	If Not rst.Eof Then
		oblog.ShowMsg "�õȼ����Ѿ�����",""
		Exit Sub
	End If
	Set rst=conn.Execute("select groupid From oblog_groups Where g_level=" & sLevel & tsql )
	If Not rst.Eof Then
		oblog.ShowMsg "�õȼ�����Ѿ�����",""
		Exit Sub
	End If
	rst.Close
	Dim nAddMode
	nAddMode=false
	Set rst=Server.CreateObject("Adodb.Recordset")
	rst.Open sql,conn,1,3
	If  rst.Eof Then
		nAddMode=true
		rst.AddNew
	End If
	rst("g_name")=sName
	rst("g_level")=sLevel
'	rst("g_style") = sStyle
	rst("g_style") = sTeamSkins
	rst("g_points")=sPoints
	rst("g_autoupdate") = sAutoUpdate
	rst("g_obcodes") = sOBCodes
	rst("g_post_day") = sDayPosts
	rst("g_post_chk") = sPostChk
	rst("g_seconddomain")=sSecondDomain
	rst("g_domain") = sDomain
	rst("g_classes") = sClasses
	rst("g_skins") = sSkins
'	rst("g_team_skins")=sTeamSkins
	rst("g_skin_edit") = sSkinEdit
	rst("g_skin_script") = sSkinScript
	rst("g_post_script") = sPostScript
	rst("g_ad_sys") = sADSystem
	rst("g_ad_user") = sADuser
	rst("g_qq_number") = sQQNumber
	rst("g_mod_note") = sModNote
	rst("g_mod_address") = sModAddress
	rst("g_mod_zhai") = sModZhai
	rst("g_mod_argue") = sModArgue
	rst("g_mod_meet") = sModActions
	rst("g_up_types") = sUpTypes
	rst("g_up_size") = sUpSize
	rst("g_up_space") = sUpSpace
	rst("g_up_watermark") = sUpWaterMark
	rst("g_is_password") = sIsPassword
	rst("g_pm_numbers")=sPMs
	rst("oneday_update")=sUpdates
	rst("is_code_addblog")=sCodePost
	rst("In_group")=sIn_group
	If oblog.cacheconfig(51)="1" Then
		rst("g_mailpost")=sMailPost
	End If
	rst("g_download") = sDownLoad
	rst("g_getpwd") = sGetPwd
	rst.Update
	If Err Then
		If nAddMode Then rst.Delete
		%>
		<script lanaguage="javascript">
			alert("������ȷ��дÿһ����Ŀ!")
			history.back();
		</script>
		<%
		Response.End
	End If
	EventLog "�����û������ϵ���ӣ��޸ģ�������Ŀ���û���IDΪ��"&OB_IIF(sGid,"��"),oblog.NowUrl&"?"&Request.QueryString
	'Response.Write "�û��ȼ� " & sName & " �����ɹ�!"
 	Response.Redirect "admin_groups.asp"
 End Sub

 Function ShowMove()
 	Dim sSelect,rs
 	Set rs=oblog.Execute("select groupid,g_name From oblog_groups")
 	Do While Not rs.Eof
 		sSelect=sSelect & "<option value="&rs(0)&">" & rs(1) & "</option>" & vbcrlf
 		rs.MoveNext
	Loop
 	Set rs=Nothing
 %>
 <script language="javascript">
 	function checksubmit1(){
 		if(document.form1.g1.value==document.form1.g2.value){
 			alert("������ͬһ����������ת��!")
 			return false;
 			}
}
 </script>
 <div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�û��ȼ�����ת��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
 <form method="post" action="admin_groups.asp?action=savemove" id="form1" name="form1" onSubmit="return checksubmit1();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
     <td width="100%" height="25" >ѡ��Ҫ����ת�Ƶ���:
     	<select name=g1><%=sSelect%></select>&nbsp;&nbsp;&nbsp;&nbsp;Ŀ����:<select name=g2><%=sSelect%></select>&nbsp;&nbsp;&nbsp;&nbsp;<input type=submit value="ִ��">
     	</td>
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
 End Function

 Function SaveMove()
 	Dim g1,g2,rs,rst
 	g1=Int(Request.Form("g1"))
 	g2=Int(Request.Form("g2"))
 	If g1=g2 Then Response.Redirect "admin_groups.asp"
 	oblog.Execute "Update oblog_user Set user_group=" & g2 & " Where user_group=" & g1
	oblog.CountGroupUser
	EventLog "�����û����ת�Ʋ������û���ID"&g1&"ת�Ƶ��û���ID"&g2,oblog.NowUrl&"?"&Request.QueryString
	%>
	<script language="javascript">
 			alert("����ת�����!")
			document.location.href="admin_groups.asp";
 </script>
	<%
 End Function
 Set oblog = Nothing
%>
