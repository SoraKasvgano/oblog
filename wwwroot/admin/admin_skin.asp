<!--#include file="inc/inc_sys.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--系统模板管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.form.chkAll.checked){
	document.form.chkAll.checked = document.form.chkAll.checked&0;
    }
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
</SCRIPT>
<%
dim action
Dim tableName,stype
tableName = "[oblog_userskin]"
stype = "user"
Action=Trim(Request("Action"))

select case Action
	case "outuser"
		call outuser()
	case "outuserok"
		call outuserok()
	case "inuser1"
		call inuser1()
	case "inuser2"
		call inuser2()
	case "inuserok"
		call inuserok()
	case "outsys"
		call outsys()
	case "outsysok"
		call outsysok()
	case "insys1"
		call insys1()
	case "insys2"
		call insys2()
	case "insysok"
		call insysok()
	case "outteam"
		tableName = "[oblog_teamskin]"
		stype = "team"
		call outuser()
	case "outteamok"
		tableName = "[oblog_teamskin]"
		stype = "team"
		call outuserok()
	case "inteam1"
		tableName = "[oblog_teamskin]"
		stype = "team"
		call inuser1()
	case "inteam2"
		tableName = "[oblog_teamskin]"
		stype = "team"
		call inuser2()
	case "inteamok"
		tableName = "[oblog_teamskin]"
		stype = "team"
		call inuserok()
end select

sub outuserok()
	dim mdbname,rs,connskin,fso
	dim skinid,i,rsout
	mdbname=Trim(Request("mdbname"))
	set fso=Server.CreateObject(oblog.CacheCompont(1))
	if fso.FileExists(Server.MapPath(mdbname))=False then
		Response.Write("<script language=javascript>alert('“"&mdbname&"”不存在！');history.back();</script>")
		Response.End
	end if
	Set connskin = Server.CreateObject("ADODB.Connection")
	connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
	set rsout=Server.CreateObject("adodb.recordset")
	rsout.open "select * from skin",connskin,2,3
	skinid=split(Request("id"))
	for i=0 to ubound(skinid)
		set rs=oblog.execute("select * from "&tableName&" where id="&CLng(skinid(i)))
		rsout.addnew
		rsout("type")=stype
		rsout("skinname")=rs("userskinname")
		rsout("skinmain")=rs("skinmain")
		rsout("skinshowlog")=rs("skinshowlog")
		rsout("skinauthor")=rs("skinauthor")
		rsout("skinauthorurl")=rs("skinauthorurl")
		rsout("skinpic")=rs("skinpic")
		rsout.update
	next
	rsout.close
	set rsout=nothing
	set rs=Nothing
	EventLog "进行导出用户（群组）模板的操作，目标模板ID："&Join(skinid)&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Write("导出成功！")
end sub

sub inuserok()
	dim mdbname,rs,connskin
	dim skinid,i,rsin
	mdbname=Trim(Request("mdbname"))
	Set connskin = Server.CreateObject("ADODB.Connection")
	connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
	set rsin=Server.CreateObject("adodb.recordset")
	if not IsObject(conn) then link_database
	rsin.open "select top 1 * from "&tableName&"",conn,2,3
	skinid=split(Request("id"))
	for i=0 to ubound(skinid)
		set rs=connskin.execute("select * from skin where type='"&stype&"' and id="&CLng(skinid(i)))
		rsin.addnew
		rsin("userskinname")=rs("skinname")
		rsin("skinmain")=rs("skinmain")
		rsin("skinshowlog")=rs("skinshowlog")
		rsin("skinauthor")=rs("skinauthor")
		rsin("skinauthorurl")=rs("skinauthorurl")
		rsin("skinpic")=rs("skinpic")
		rsin("isdefault") = 0
		rsin("ispass") = 1
		If stype = "user" Then
			rsin("classid") = 0
		Else
			rsin("classid") = 2
		End if
		rsin.update
	next
	rsin.close
	set rsin=nothing
	set rs=Nothing
	EventLog "进行导入用户（群组）模板的操作，目标模板ID："&Join(skinid)&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Write("导入成功！")
end sub

sub outsysok()
	dim mdbname,rs,connskin,fso
	dim skinid,i,rsout
	mdbname=Trim(Request("mdbname"))
	set fso=Server.CreateObject(oblog.CacheCompont(1))
    if (fso.FileExists(Server.MapPath(mdbname))) Then
		Set connskin = Server.CreateObject("ADODB.Connection")
		connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
		set rsout=Server.CreateObject("adodb.recordset")
		rsout.open "select top 1 * from skin",connskin,2,3
		skinid=split(Request("id"))
		for i=0 to ubound(skinid)
			set rs=oblog.execute("select * from oblog_sysskin where id="&CLng(skinid(i)))
			rsout.addnew
			rsout("type")="sys"
			rsout("skinname")=rs("sysskinname")
			rsout("skinmain")=rs("skinmain")
			rsout("skinshowlog")=rs("skinshowlog")
			rsout("skinauthor")=rs("skinauthor")
			rsout.update
		next
		rsout.close
		set rsout=nothing
		set rs=Nothing
		EventLog "进行导出系统模板的操作，目标模板ID："&Join(skinid)&"",oblog.NowUrl&"?"&Request.QueryString
		Response.Write("导出成功！")
	Else
		Response.Write("<script language=javascript>alert('“"&mdbname&"”不存在！');history.back();</script>")
		Response.End
	End if
end sub

sub insysok()
	dim mdbname,rs,connskin
	dim skinid,i,rsin
	mdbname=Trim(Request("mdbname"))
	Set connskin = Server.CreateObject("ADODB.Connection")
	connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
	set rsin=Server.CreateObject("adodb.recordset")
	if not IsObject(conn) then link_database
	rsin.open "select * from oblog_sysskin",conn,2,3
	skinid=split(Request("id"))
	for i=0 to ubound(skinid)
		set rs=connskin.execute("select * from skin where type='sys' and id="&CLng(skinid(i)))
		rsin.addnew
		rsin("sysskinname")=rs("skinname")
		rsin("skinmain")=rs("skinmain")
		rsin("skinshowlog")=rs("skinshowlog")
		rsin("skinauthor")=rs("skinauthor")
		rsin.update
	next
	rsin.close
	set rsin=nothing
	set rs=Nothing
	EventLog "进行导入系统模板的操作，目标模板ID："&Join(skinid)&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Write("导入成功！")
end sub
%>
<%
sub outuser()
dim rs,Temp
set rs=oblog.execute("select * from "&tableName&" ")
If tableName = "[oblog_teamskin]" Then
	Temp = "outteamok"
Else
	Temp = "outuserok"
End if
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">模板导出</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form name="form" method="post" action="admin_skin.asp">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td width="63" align="center" >选中</td>
      <td width="120"> <div align="center">ID</div></td>
      <td width="288" > <div align="center">名称</div></td>
      <td width="292" > <div align="center">作者</div></td>
    </tr>
    <%
while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="63" align="center"><input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("id"))%>'></td>
      <td width="120"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="288" > <div align="center"><%= rs("userskinname") %></div></td>
      <td > <div align="center"><%= rs("skinauthor") %></div></td>
    </tr>
    <%
rs.movenext
wend
%>
    <tr>
      <td height="40" colspan="4" align="center" class="tdbg"> <div align="left">
			<input type="hidden" id="action" name="action" value="<%=Temp%>" />
          <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
          选中所有模板　导出数据库名：
          <input name="mdbname" type="text" id="mdbname" value="../skin/skin.mdb" size="20" maxlength="50">
          <input type="submit" name="Submit" value=" 导出模板 ">
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
<%end sub
sub inuser1()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">模板导入</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
		<%If Action ="inuser1" Then %>
<form name="form" method="post" action="admin_skin.asp?action=inuser2">
<%Else%>
<form name="form" method="post" action="admin_skin.asp?action=inteam2">
<%End if%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
     <tr>
      <td width="763" height="40" align="center" class="tdbg"> <div align="left">
          　导入数据库名：
          <input name="mdbname" type="text" id="mdbname" value="../skin/skin.mdb" size="20" maxlength="50">
          <input type="submit" name="Submit" value=" 下一步 ">
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
<%end sub

sub inuser2()
dim connskin,rs,mdbname,fso
mdbname=Trim(Request("mdbname"))
	set fso=Server.CreateObject(oblog.CacheCompont(1))
	if fso.FileExists(Server.MapPath(mdbname))=False then
		Response.Write("<script language=javascript>alert('“"&mdbname&"”不存在！');history.back();</script>")
		Response.End
	end if
Set connskin = Server.CreateObject("ADODB.Connection")
connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
set rs=connskin.execute("select * from skin where type='"&stype&"'")
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">用户模板导入</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
		<% If stype = "user" Then %>
<form name="form" method="post" action="admin_skin.asp?action=inuserok">
<%Else%>
<form name="form" method="post" action="admin_skin.asp?action=inteamok">
<%End if%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td width="63" align="center" >选中</td>
      <td width="120"> <div align="center">ID</div></td>
      <td width="288" > <div align="center">名称</div></td>
      <td width="292" > <div align="center">作者</div></td>
    </tr>
    <%
while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="63" align="center"><input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("id"))%>'></td>
      <td width="120"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="288" > <div align="center"><%= rs("skinname") %></div></td>
      <td > <div align="center"><%= rs("skinauthor") %></div></td>
    </tr>
    <%
rs.movenext
wend
%>
    <tr>
      <td height="40" colspan="4" align="center" class="tdbg"> <div align="left">
          <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
          选中所有模板
          <input type="submit" name="Submit" value=" 导入模板 ">
          <input type="hidden" name="mdbname" value="<%=mdbname%>">
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
<%end sub%>
<%
sub outsys()
dim rs
set rs=oblog.execute("select * from oblog_sysskin ")
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">系统模板导出</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form name="form" method="post" action="admin_skin.asp?action=outsysok">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td width="63" align="center" >选中</td>
      <td width="120"> <div align="center">ID</div></td>
      <td width="288" > <div align="center">名称</div></td>
      <td width="292" > <div align="center">作者</div></td>
    </tr>
    <%
while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="63" align="center"><input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("id"))%>'></td>
      <td width="120"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="288" > <div align="center"><%= rs("sysskinname") %></div></td>
      <td > <div align="center"><%= rs("skinauthor") %></div></td>
    </tr>
    <%
rs.movenext
wend
%>
    <tr>
      <td height="40" colspan="4" align="center" class="tdbg"> <div align="left">
          <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
          选中所有模板　导出数据库名：
          <input name="mdbname" type="text" id="mdbname" value="../skin/skin.mdb" size="20" maxlength="50">
          <input type="submit" name="Submit" value=" 导出模板 ">
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
<%end sub

sub insys1()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">系统模板导入</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form name="form" method="post" action="admin_skin.asp?action=insys2">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
     <tr>
      <td width="763" height="40" align="center" class="tdbg"> <div align="left">
          　导入数据库名：
          <input name="mdbname" type="text" id="mdbname" value="../skin/skin.mdb" size="20" maxlength="50">
          <input type="submit" name="Submit" value=" 下一步 ">
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
<%end sub

sub insys2()
dim connskin,rs,mdbname,fso
mdbname=Trim(Request("mdbname"))
set fso=Server.CreateObject(oblog.CacheCompont(1))
    if fso.FileExists(Server.MapPath(mdbname))=False then
    Response.Write("<script language=javascript>alert('“"&mdbname&"”不存在！');history.back();</script>")
    Response.End
    end if
Set connskin = Server.CreateObject("ADODB.Connection")
connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
set rs=connskin.execute("select * from skin where type='sys'")
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">系统模板导入</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form name="form" method="post" action="admin_skin.asp?action=insysok">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td width="63" align="center" >选中</td>
      <td width="120"> <div align="center">ID</div></td>
      <td width="288" > <div align="center">名称</div></td>
      <td width="292" > <div align="center">作者</div></td>
    </tr>
    <%
while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="63" align="center"><input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("id"))%>'></td>
      <td width="120"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="288" > <div align="center"><%= rs("skinname") %></div></td>
      <td > <div align="center"><%= rs("skinauthor") %></div></td>
    </tr>
    <%
rs.movenext
wend
%>
    <tr>
      <td height="40" colspan="4" align="center" class="tdbg"> <div align="left">
          <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
          选中所有模板
          <input type="submit" name="Submit" value=" 导入模板 ">
          <input type="hidden" name="mdbname" value="<%=mdbname%>">
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
<%end Sub
Set oblog = Nothing
%>
</body>
</html>