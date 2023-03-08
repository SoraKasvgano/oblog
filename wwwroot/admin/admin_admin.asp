<!--#include file="inc/inc_sys.asp"-->
<%
dim rs, rst,sql,roleId
dim Action,iCount,adminname,strPara
strPara=LCase(Request.QueryString)
Action=Trim(Request("Action"))
adminname=session("adminname")
CheckSafePath(0)
Set rst=Server.CreateObject("Adodb.Recordset")
Set rst=oblog.Execute("select roleid,r_name From oblog_roles Order By roleid")
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
<script language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    }
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled!=true)
       e.checked = form.chkAll.checked;
    }
}

function CheckAdd()
{
  if(document.form1.username.value=="")
    {
      alert("用户名不能为空！");
	  document.form1.username.focus();
      return false;
    }

  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }

  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();
      return false;
    }
/*   if (document.form1.Purview[1].checked==true){
	GetClassPurview();
  }
  */
}
function CheckModifyPwd()
{
  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();
      return false;
    }
}

</script>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">管 理 员 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg">
    <td height="22" colspan="2" align="center"><strong>管 理 员 管 理</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="admin_admin.asp">管理员管理首页</a>&nbsp;|&nbsp;<a href="admin_admin.asp?Action=Add">新增管理员</a></td>
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
select Case Lcase(Action)
	Case "add"
		call AddAdmin()
	Case "saveadd"
		If CheckSafePath(0) Then call SaveAdd()
	Case "edit"
		Call EditAdmin()
	Case "saveedit"
		Call SaveEdit
	Case "del"
		If CheckSafePath(0) Then call DelAdmin()
	Case Else
		call main()
end select


Sub main()
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from oblog_admin order by roleid"
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
%>
<style>

tr td {
padding:5px 0!important;
}

</style>

<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">管 理 员 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="myform" method="Post" action="admin_admin.asp" onSubmit="return confirm('确定要删除选中的管理员吗？');">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr align="center" class="title">
            <td width="32"><font style="color:#000;font-weight:600;">选中</font></td>
            <td width="40"><font style="color:#000;font-weight:600;">序号</font></td>
            <td width="180"><font style="color:#000;font-weight:600;">权限角色</font></td>
            <td width="100"><font style="color:#000;font-weight:600;">用 户 名</font></td>
<!--             <td width="80"><font style="color:#000;font-weight:600;">绑定博客</font></td> -->
            <td width="100"><font style="color:#000;font-weight:600;">最后登录IP</font></td>
            <td><font style="color:#000;font-weight:600;">最后登录时间</font></td>
            <td width="60"><font style="color:#000;font-weight:600;">登录次数</font></td>
            <td width="32"><font style="color:#000;font-weight:600;">修改</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr align="center" class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
            <td><input name="ID" type="checkbox" id="ID" value="<%=rs("ID")%>" <%if rs("UserName")=AdminName then Response.write " disabled"%> onClick="unselectall()"></td>
            <td><%=rs("ID")%></td>
            <td><%
            	If Not IsNull(rs("roleid")) Then
	            	rst.Filter="roleid=" & rs("roleid")
	            	If Not rst.Eof Then
	            		Response.Write rst("r_name")
	            	Else
						If rs("roleid") = 0 Then
	            			Response.Write "<font color=green>系统管理员</font>"
						Else
							Response.Write "<font color=blue>无权限管理员组</font>"
						End if
	            	End If
	            Else
	            	Response.Write "<font color=green>系统管理员</font>"
	            End If
            	%></td>
            <td>
              <%
				if rs("username")=AdminName then
					Response.write "<font color=red><b>" & rs("UserName") & "</b></font>"
				else
					Response.write rs("UserName")
				end if
				%>
            </td>
<!--             <td>
              <%
				if rs("userid")<>""  then
					Response.write "<a href=""../go.asp?userid=" & rs("userid") & """ target=_blank>" & rs("userid") & "</a>"
				else
					Response.write "&nbsp;"
				end if
				%>
            </td> -->
            <td>
              <%
				if rs("LastLoginIP")<>"" then
					Response.write rs("LastLoginIP")
				else
					Response.write "&nbsp;"
				end if
				%>
            </td>
            <td>
              <%
				if rs("LastLoginTime")<>"" then
					Response.write rs("LastLoginTime")
				else
					Response.write "&nbsp;"
				end if
				%>
            </td>
            <td>
			<%
			    If Not IsNull(rs("LoginTimes")) Then
					If rs("LoginTimes")<>"" Then
						Response.write rs("LoginTimes")
					Else
						Response.write 0
						oblog.execute ("update [oblog_admin]  set LoginTimes=0 where id="&uid)
					End If
				Else
					oblog.execute ("update [oblog_admin]  set LoginTimes=0")
				End if
				%>
            </td>
            <td>
            	<%If rs("roleid")>0 Then%>
            		<a href="admin_admin.asp?action=edit&id=<%=rs("id")%>">修改</a>
            	<%Else%>
<!--             <s>修改</s> -->
			<a href="admin_admin.asp?action=edit&id=<%=rs("id")%>">修改</a>
            		<%End If%>
            	</td>
          </tr>
          <%
	rs.MoveNext
loop
  %>
          <tr class="title">
            <td colspan="9"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有管理员<input name="Action" type="hidden" id="Action" value="Del">
              <input name="Submit" type="submit" id="Submit" value="删除选中的管理员"></td>
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
	rs.Close
	set rs=Nothing
end sub

sub AddAdmin()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">新 增 管 理 员</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="post" action="admin_admin.asp" name="form1" onSubmit="javascript:return CheckAdd();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border" >
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">选择管理员角色：</div></td>
      <td width="65%" class="tdbg">
      	<select name="roleid">
      		<option value="0">系统管理员(不可绑定用户ID)</option>
      		<%
      		If Not rst.Eof Then
	      		rst.Movefirst
	      		Do While Not rst.Eof
	      			%>
	      		<option value="<%=rst(0)%>"><%=rst(1)%></option>
	      			<%
	      			rst.MoveNext
	      		Loop
	      	End If
      		%>
    	</select>
      	</td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">用 户 名：</div></td>
      <td width="65%" class="tdbg"><input name="username" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">初始密码： </div></td>
      <td width="65%" class="tdbg"><font size="2">
        <input type="password" name="Password">
        </font></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">确认密码：</div></td>
      <td width="65%" class="tdbg"><font size="2">
        <input type="password" name="PwdConfirm">
        </font></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">绑定用户数字ID：</div></td>
      <td width="65%" class="tdbg"><input name="userid" type="text">(前台管理员可绑定用户ID，但是仅可以绑定一个) &nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td colspan="2"><div align="center">
          <input name="Action" type="hidden" id="Action" value="SaveAdd">
          <input  type="submit" name="Submit" value=" 添 加 " style="cursor:hand;">
          &nbsp;
          <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;">
        </div></td>
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
end sub
%>
<%
sub SaveAdd()
	dim username, password,PwdConfirm,userid
	If Instr(strPara,"password") Then
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""注意：外部恶意连接！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		Exit Sub
	End If
	username=Trim(Request("username"))
	password=Trim(Request("Password"))
	roleid=Trim(Request("roleid"))
	userid=Trim(Request("userid"))
	If roleId="" Then
		roleId=0
	Else
		roleid=Int(roleid)
	End If
	If userid="" Then
		userid=0
	Else
		if not IsNumeric(userid) then
			oblog.ShowMsg "用户id必须为数字！","back"
			exit sub
		end if
		userid=Int(userid)
	End If
	sql="select * from oblog_admin where username='"&username&"'"
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if not (rs.bof and rs.EOF) then
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""对不起！此用管理员已经存在，请更换用户名再进行注册！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		rs.close
		set rs=nothing
		exit sub
	end if
   	rs.addnew
   	rs("roleid")=roleid
 	rs("username")=username
   	rs("password")=md5(password)
   	rs("userid")=userid
	rs.update
    rs.Close
	set rs=Nothing
	If userid>0 Then oblog.Execute "Update Oblog_user Set roleid=" & roleid & " Where userid=" & userid
	Call main()
end sub

sub SaveEdit()
	dim id,username, password,PwdConfirm,userid,userid1
	If Instr(strPara,"password") Then
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""注意：外部恶意连接！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		Exit Sub
	End If
	id=Trim(Request("id"))
	username=Trim(Request("username"))
	password=Trim(Request("Password"))
	PwdConfirm=Trim(Request("PwdConfirm"))
	roleid=Trim(Request("roleid"))
	userid=Trim(Request("userid"))
	if password<>PwdConfirm then
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""两次输入的密码不同，如果不修改请留空！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		exit sub
	end If
	If password <> "" Then
		if Len(password)<8 then
			Response.Write"<script language=JavaScript>"
			Response.Write"alert(""密码长度最少为8位，如果不修改请留空！"");"
			Response.Write"window.history.go(-1);"
			Response.Write"</script>"
			exit sub
		end If
	End if
	If roleId="" Then
		roleId=0
	Else
		roleid=Int(roleid)
	End If
	If userid="" Then
		userid=0
	Else
		userid=Int(userid)
	End If
	sql="select * from oblog_admin where id=" & id
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	If Not IsNull(rs("userid")) Then userid1=rs("userid")
   	rs("roleid")=roleid
   	If password<>"" Then rs("password")=md5(password)
   	rs("userid")=userid
	rs.update
    rs.Close
	set rs=Nothing
	'先取消
	If userid1<>"" Then oblog.Execute "Update Oblog_user Set roleid=0 Where userid=" & userid1
	'重新绑定
	If userid>0 Then oblog.Execute "Update Oblog_user Set roleid=" & roleid & " Where userid=" & userid
	EventLog "进行了添加(修改)管理员的操作，目标管理员ID为："&OB_IIF(id,"无"),oblog.NowUrl&"?"&Request.QueryString
	Call main()
end sub

sub DelAdmin()
	dim UserID
	If Instr(strPara,"del") Then
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""注意：外部恶意连接！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		Exit Sub
	End If
	UserID=Trim(Request("ID"))
	if UserID="" then
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""请选择要删除的管理员！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=Replace(UserID," ","")
		sql="select * from oblog_Admin where ID in (" & UserID & ")"
	else
		UserID=CLng(UserID)
		sql="select * from oblog_Admin where ID=" & UserID
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	do while not rs.eof
		rs.delete
		rs.update
		rs.movenext
	loop
	rs.close
	set rs=Nothing
	EventLog "进行了删除管理员的操作，目标管理员ID为："&UserID,oblog.NowUrl&"?"&Request.QueryString
	call main()
end sub

sub EditAdmin()
Dim adminid
adminid=clng(Request("id"))
'Set rs=oblog.Execute("select * From oblog_admin Where roleid>0 And id=" & adminId)
Set rs=oblog.Execute("select * From oblog_admin Where id=" & adminId)
If rs.Eof Then
	Response.Redirect "admin_admin.asp"
End If
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">修 改 管 理 员</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="post" action="admin_admin.asp" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border" >
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">选择管理员角色：</div></td>
      <td width="65%" class="tdbg">
      	<select name="roleid" <%if rs("username")=AdminName Then Response.Write "disabled" End if%>>
			<option value="-1">请选择管理权限</option>
      		<option value="0" <%If rs("roleid")=0 Then Response.Write "selected" End if%> >系统管理员(不可绑定用户ID)</option>
      		<%
      		rst.Movefirst
      		Do While Not rst.Eof
      			%>
      		<option value="<%=rst(0)%>" <%If rst(0)=rs("roleid") Then Response.Write "selected" End if%>><%=rst(1)%></option>
      			<%
      			rst.MoveNext
      		Loop
      		%>
    	</select>
      	</td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">用 户 名：</div></td>
      <td width="65%" class="tdbg"><input name="username" type="text" value="<%=rs("username")%>" disabled> &nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">初始密码： </div></td>
      <td width="65%" class="tdbg"><font size="2">
        <input type="password" name="Password">
        </font>(不修改密码请留空)</td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">确认密码：</div></td>
      <td width="65%" class="tdbg"><font size="2">
        <input type="password" name="PwdConfirm">
        </font></td>
    </tr>
     <tr class="tdbg">
      <td width="35%" class="tdbg"><div align="right">绑定用户ID：</div></td>
      <td width="65%" class="tdbg"><input name="userid" type="text" value="<%=rs("userid")%>">(前台管理员可绑定用户ID，但是仅可以绑定一个) &nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td colspan="2"><div align="center">
      	 <input name="id" type="hidden" id="id" value="<%=rs("id")%>">
          <input name="Action" type="hidden" id="Action" value="SaveEdit">
          <input  type="submit" name="Submit" value=" 修 改 " style="cursor:hand;">
          &nbsp;
          <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;">
        </div></td>
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
</body>
</html>
<%
end Sub
Set oblog = Nothing
%>