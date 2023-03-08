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
<title>用户等级配置</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">用户等级管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="tdbg">
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30">
    	<a href="admin_groups.asp">当前组列表</a>&nbsp;|&nbsp;
    	<a href="admin_groups.asp?action=add">新增用户等级</a>&nbsp;|&nbsp;
    	<a href="admin_groups.asp?action=move">不同等级间用户转移</a>
    </td>
  </tr>
  <tr class="tdbg">
  <td height="30" colspan="2" >
    	<font color="red"><strong>等级级别最小的组被默认为用户注册后所属的默认等级</strong></font>
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
				ErrMsg="该等级中有"&rst(0)&"用户存在,请先将该部分用户转移到其他等级,此处列出除该等级外的等级列表,进行选择"

			Else
				conn.execute("Delete From oblog_groups Where groupid=" & sGid)
				ErrMsg="该等级已经被删除"
			End If
		End If
		EventLog "进行用户组的删除操作，目标用户组ID为"&sGid,oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "删除成功","admin_groups.asp"
	Case "restat"
		Set rst = oblog.execute ("select count(userid) from oblog_user where user_group = "&sGid)
		oblog.execute ("update oblog_groups set g_members ="&rst(0)&" where groupid = " & sGid)
		rst.close
		Set rst=Nothing
		oblog.ShowMsg "重新统计成功","admin_groups.asp"
	Case  Else
		Call ShowList
End select

conn.Close
Set conn=Nothing
Dim sJs
sJs=JsValid("form1","g_name","1",1,50,"必须输入等级名称!")
sJs=sJs & JsValid("form1","g_uptypes","1",1,100,"必须填写上传文件类型！")
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
		<li class="main_top_left left">用户等级信息</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
	<%
	If rst.Eof Then
	%>
		<tr><td align="center">目前还没有定义任何用户等级信息</td></tr>
	<%
		rst.Close
		Set rst=Nothing
		Exit Sub
	End If
	%>
	<tr><td>等级级别</td><td>等级名称<td>积分限制</td><td>是否自动升级</td><td>用户数</td><td>操作</td>
	<%
	Do While Not rst.Eof
	%>
		<tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
		<td><%=rst(5)%></td>
		<td><%=rst(1)%></td>
		<td><%=rst(2)%></td>
		<td><%If rst(3)="1" Then Response.Write "是" Else Response.Write "否"%></td>
		<td><%=rst(4)%></td>
		<td><a href="admin_groups.asp?action=restat&gid=<%=rst(0)%>">重新统计</a>&nbsp;&nbsp;<a href="admin_groups.asp?action=show&gid=<%=rst(0)%>">修改</a>&nbsp;&nbsp;
		<%If rst(4)=0 Then%>
			<a href="admin_groups.asp?action=del&gid=<%=rst(0)%>" onclick="return confirm('确定要删除选中的用户组吗？');">删除</a>
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
		<li class="main_top_left left">等级信息配置</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="post" action="admin_groups.asp?action=save&gid=<%=Gid%>" id="form1" name="form1" onSubmit="return checksubmit();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
     <td width="348" height="25" >等级名称</td>
     <td height="25" ><% Call EchoInput("g_name",40,40,sName)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >等级级别(不允许重复)：</td>
      <td height="25" ><% Call EchoInput("g_level",40,40,sLevel)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >升级到该等级的最小积分：</td>
      <td height="25" ><% Call EchoInput("g_points",40,10,sPoints)%>(如果该等级不允许升级，请将其设置为0)</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > 是否允许自动升级</td>
      <td> <% Call EchoRadio("g_autoupdate","","",sAutoupdate)%>  </td>
	  </tr>
	  <!--
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >组的css样式,注意在程序中将使用 style="xxx" 的形式</td>
      <td height="25" ><% Call EchoInput("g_font",40,20,sStyle)%></td>
    </tr>
    -->
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否允许修改二级域名</td>
      <td height="25" ><% Call EchoRadio("g_seconddomain","","",sSecondDomain)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否可以绑定顶级域名</td>
      <td height="25" ><% Call EchoRadio("g_domain","","",sDomain)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >每天最多可发出的邀请数目，0为不允许发送</td>
      <td height="25" ><% Call EchoInput("g_obcodes",40,40,sOBCodes)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >发布的日志是否需要审核</td>
      <td height="25" > <% Call EchoRadio("g_logchk","","",sPostChk)%></td>
    </tr>
	 <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >允许发布的一级日志分类<br/>(不选择则允许发布在任意分类)</td>
      <td>
	  <%
	  Dim sTmp,sChecked,rstTmp,i
	  sTmp="," & sClasses & ","
	  sChecked=""
	  Set rstTmp=conn.Execute("select * From oblog_logclass Where depth=0 And idType=0")
	  If rstTmp.Eof Then
			Response.Write "还没有定义任何分类"
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
      <td width="348" height="25" >每天允许发布的日志数目</td>
      <td><% Call EchoInput("g_dayposts",40,40,sDayPosts)%>(0为不限制)</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >日志中是否允许使用脚本</td>
      <td height="25" > <% Call EchoRadio("g_post_script","","",sPostScript)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >允许使用的群组模板分类<br/>(如果为不选择则允许使用所有模板,下同)</td>
      <td height="25" >
	  <%
	  sTmp=Replace("," & sTeamSkins & ","," ","")
	  sChecked=""
	  i=0
	  Set rstTmp=conn.Execute("select * From oblog_skinclass Where iType=1")
	  If rstTmp.Eof Then
			Response.Write "还没有定义任何分类"
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
      <td width="348" height="25" >允许使用的用户模板分类</td>
      <td height="25" >
	  <%
	  sTmp=Replace("," & sSkins & ","," ","")
	  sChecked=""
	  i=0
	  Set rstTmp=conn.Execute("select * From oblog_skinclass Where iType=0")
	  If rstTmp.Eof Then
			Response.Write "还没有定义任何分类"
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
      <td width="348" height="25" >是否允许编辑模板</td>
      <td height="25" > <% Call EchoRadio("g_skin_edit","","",sSkinEdit)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >模板中是否使用脚本</td>
      <td height="25" > <% Call EchoRadio("g_skin_script","","",sSkinScript)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >允许创建的群组数目(0为不允许创建)</td>
     <td><% Call EchoInput("g_qqnumbers",40,40,sQQNumber)%></td>
    </tr>
    <!--
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否可以使用通讯录</td>
       <td height="25" > <% Call EchoRadio("g_mod_address","","",sModAddress)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否允许使用网摘</td>
       <td height="25" > <% Call EchoRadio("g_mod_Zhai","","",sModZhai)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否允许个人记事本</td>
       <td height="25" > <% Call EchoRadio("g_mod_note","","",sModNote)%></td>
    </tr>
    -->
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >上传空间的大小(0为不限制,-1为不允许)</td>
      <td height="25" ><% Call EchoInput("g_UpSpace",40,40,sUpSpace)%>KB</td>
    </tr>
	 <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >单一文件限制</td>
      <td height="25" ><% Call EchoInput("g_UpSize",40,40,sUpSize)%>KB</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >上传文件类型</td>
      <td height="25" ><% Call EchoInput("g_uptypes",50,100,sUpTypes)%>(用|分割)</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >图片是否加水印</td>
      <td height="25" > <% Call EchoRadio("g_upwatermark","","",sUpWaterMark)%> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否允许整站加密</td>
      <td height="25" > <% Call EchoRadio("g_is_password","","",sIsPassword)%> </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >用户短信息数目限制</td>
      <td height="25" ><% Call EchoInput("g_pms",5,5,sPMs)%></td>
    </tr>

    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否强制显示广告</td>
      <td height="25" ><% Call EchoRadio("g_ad_system","","",sAdSystem)%></td>
    </tr>
	    <!--
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否允许用户发布自己的广告</td>
      <td height="25" > <% Call EchoRadio("g_ad_user","","",sAdUser)%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否允许用户启用脚本(需要进行审核,每次修改后仍然需要审核)</td>
      <td height="25" > <% Call EchoRadio("g_modscript","","",sModScript)%></td>
    </tr>
      <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否允许用户发起辩论申请</td>
      <td height="25" > <% Call EchoRadio("g_modArgue","","",sModArgue)%></td>
    </tr>

	 <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否允许用户发起活动申请</td>
      <td height="25" > <% Call EchoRadio("g_modactions","","",sModActions)%></td>
    </tr>
    -->
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td height="25" >用户发表日志是否需要验证码：</td>
      <td height="25" ><% Call EchoRadio("is_code_addblog","","",sCodePost)%></td>
    </tr>
     <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >每天允许用户更新全站多少次：</td>
      <td height="25" ><% Call EchoInput("oneday_update",5,5,sUpdates)%>次(设置为0则不进行限制)</td>
    </tr>
	  <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >允许加入的群组数：</td>
      <td height="25" ><% Call EchoInput("sIn_group",5,5,sIn_group)%>个</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否允许下载附件：</td>
      <td height="25" ><% Call EchoRadio("g_download","","",sDownLoad)%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >是否允许取回密码：</td>
      <td height="25" ><% Call EchoRadio("g_getpwd","","",sGetPwd)%></td>
    </tr>
    <%If oblog.CacheConfig(51)="1" Then%>
     <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >是否允许使用邮件或彩信发布日志：</td>
      <td height="25" ><% Call EchoRadio("g_mailpost","","",sMailPost)%></td>
    </tr>
    <%End If%>
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
	'进行数据校验
	If sName="" Then oblog.AddErrStr("必须填写组名称")
  	If sLevel="" Or Not IsNumeric(sLevel)   Then oblog.AddErrStr("<li>必须填写组级别</li>")
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
		oblog.ShowMsg "该等级名已经存在",""
		Exit Sub
	End If
	Set rst=conn.Execute("select groupid From oblog_groups Where g_level=" & sLevel & tsql )
	If Not rst.Eof Then
		oblog.ShowMsg "该等级序号已经存在",""
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
			alert("必须正确填写每一个项目!")
			history.back();
		</script>
		<%
		Response.End
	End If
	EventLog "进行用户组资料的添加（修改）操作，目标用户组ID为："&OB_IIF(sGid,"无"),oblog.NowUrl&"?"&Request.QueryString
	'Response.Write "用户等级 " & sName & " 操作成功!"
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
 			alert("不能在同一组间进行数据转移!")
 			return false;
 			}
}
 </script>
 <div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">用户等级数据转移</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
 <form method="post" action="admin_groups.asp?action=savemove" id="form1" name="form1" onSubmit="return checksubmit1();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
     <td width="100%" height="25" >选择要进行转移的组:
     	<select name=g1><%=sSelect%></select>&nbsp;&nbsp;&nbsp;&nbsp;目标组:<select name=g2><%=sSelect%></select>&nbsp;&nbsp;&nbsp;&nbsp;<input type=submit value="执行">
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
	EventLog "进行用户组的转移操作，用户组ID"&g1&"转移到用户组ID"&g2,oblog.NowUrl&"?"&Request.QueryString
	%>
	<script language="javascript">
 			alert("数据转移完成!")
			document.location.href="admin_groups.asp";
 </script>
	<%
 End Function
 Set oblog = Nothing
%>
