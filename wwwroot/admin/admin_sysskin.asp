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
<%
dim action
Action=Trim(Request("Action"))

select case Action
	case "saveconfig"
		call saveconfig()
	case "showskin"
		call showskin()
	case "modiskin"
		call modiskin()
	case "savedefault"
		call savedefault()
	case "delconfig"
		call delconfig()
	case "addskin"
		call addskin()
	case "saveaddskin"
		call saveaddskin()
end select

sub showskin()
dim rs
set rs=oblog.execute("select * from oblog_sysskin ")
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">系统模板管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form name="form2" method="post" action="admin_sysskin.asp?action=savedefault">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
  <tr class="topbg">
      <td height="25" >
<div align="center">ID</div></td>
    <td> <div align="center">名称</div></td>
    <td><div align="center">作者</div></td>
      <td>
<div align="center">默认模板</div></td>
      <td>
        <div align="center">模板管理</div></td>
  </tr>
 <%
while not rs.eof
%>
          <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
          <td width="25">
              <div align="center"><%= rs("id") %>&nbsp;</div></td>
          <td width="120" > <div align="center"><div align="center"><%If rs("isdefault")=1 Then Response.Write "<font color=red>" & rs("sysskinname") & "</font>" Else Response.Write rs("sysskinname") End If %></div></td>
          <td width="100" ><div align="center"><%= rs("skinauthor") %></div></td>
            <td width="50">              <div align="center">
                <input name="radiobutton" type="radio" class="tdbg" value='<%=rs("id")%>' <%if rs("isdefault")=1 then Response.Write "checked" %>>
            </div></td>
            <td>
                <div align="center">
				<a href="../admin_edit.asp?action=modiskin&skintype=sys&t=0&editm=1&skinorder=0&id=<%=rs("id")%>"  target="_blank">修改主模板</a>
                        　<a href="../admin_edit.asp?action=modiskin&skintype=sys&t=0&editm=1&skinorder=1&id=<%=rs("id")%>" target="_blank">修改副模板</a>
				<a href="admin_sysskin.asp?action=modiskin&id=<%=rs("id")%>">修改模板(文本方式)</a>
          　<a href="admin_sysskin.asp?action=delconfig&id=<%=rs("id")%>" onclick=return(confirm("确定要删除这个模板吗？"))>删除模板</a></div></td>
        </tr>

      <%
rs.movenext
wend
%>
    <tr>
    <td height="40" colspan="5" align="center" class="tdbg"> <div align="center">
          <input type="submit" name="Submit" value="保存设置">
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
	set rs=nothing
end sub

sub saveconfig()
	dim rs,sql
	if Trim(Request("sysskinname"))="" then oblog.sys_err("模板名不能为空"):Response.End()
	if Trim(Request("skinmain"))="" then oblog.sys_err("主模板不能为空"):Response.End()
	if Trim(Request("skinshowlog"))="" then oblog.sys_err("副模板不能为空"):Response.End()
	set rs=Server.CreateObject("adodb.recordset")
	sql="select * from oblog_sysskin where id="&CLng(Request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs("sysskinname")=Trim(Request("sysskinname"))
	rs("skinauthor")=Trim(Request("skinauthor"))
	rs("skinmain")=Request("skinmain")
	rs("skinshowlog")=Request("skinshowlog")
	rs.update
	rs.close
	set rs=nothing
	oblog.reloadsetup
	EventLog "进行了修改系统模板操作（文本方式），目标模板ID："&Request.QueryString("id")&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "修改成功",""
end sub
sub savedefault()
	dim isdefaultID
	isdefaultID=CLng(Trim(Request("radiobutton")))
	oblog.execute("update oblog_sysskin set isdefault=0")
	oblog.execute("update oblog_sysskin set isdefault=1 where id="&isdefaultID)
	EventLog "进行了设定系统默认模板操作，目标模板ID："&isdefaultID&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""修改成功！"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub delconfig()
	oblog.execute("delete from oblog_sysskin where id="&CLng(Request.QueryString("id")))
	EventLog "进行了删除系统模板操作，目标模板ID："&Request.QueryString("id")&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Redirect "admin_sysskin.asp?action=showskin"
end sub

sub saveaddskin()
	dim rs,sql
	set rs=Server.CreateObject("adodb.recordset")
	if Trim(Request("sysskinname"))="" then oblog.sys_err("模板名不能为空"):Response.End()
	if Trim(Request("skinmain"))="" then oblog.sys_err("主模板不能为空"):Response.End()
	if Trim(Request("skinshowlog"))="" then oblog.sys_err("副模板不能为空"):Response.End()
	sql="select * from oblog_sysskin where id="&CLng(Request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs.addnew
	rs("sysskinname")=Trim(Request("sysskinname"))
	rs("skinauthor")=Trim(Request("skinauthor"))
	rs("skinmain")=Trim(Request("skinmain"))
	rs("skinshowlog")=Trim(Request("skinshowlog"))
	rs.update
	rs.close
	set rs=Nothing
	EventLog "进行了添加系统模板操作",oblog.NowUrl&"?"&Request.QueryString
	Response.Redirect "admin_sysskin.asp?action=showskin"
end sub

sub modiskin()
	dim rs
	set rs=oblog.execute("select * from oblog_sysskin where id="&CLng(Request.QueryString("id")))
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">修改系统模板</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
    <tr>
    <td height="22" colspan=2 class="topbg"><strong>修改系统模板</strong></td>
    </tr>
    <tr class="tdbg">

    <td width="253" height="30"><strong>现在修改的模板是：<%=rs("sysskinname")%></strong></td>

    <td width="516" height="30"><a href="admin_sysskin.asp?action=modiskin&id=<%=rs("id")%>">修改模板</a>
　　<a href="admin_sysskin.asp?action=showskin">返回管理菜单</a>
      <a href="admin_skin_help.asp" target="_blank"><strong>模板标记帮助</strong></a>
	 </td>
    </tr>
</table>
<br />
<form method="POST" action="admin_sysskin.asp?id=<%=CLng(Request.QueryString("id"))%>" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="769" height="22" class="topbg"><strong>模板参数</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">模板名称：
        <input name="sysskinname" type="text" id="sysskinname" value=<%=rs("sysskinname")%>>
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <strong>主模板：</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"><%if rs("skinmain")<>"" then Response.Write Server.HtmlEncode(rs("skinmain")) else Response.Write("")%></textarea>
        <br>
        <br>
        <strong>副模板： <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"><%if rs("skinshowlog")<>"" then Response.Write Server.HtmlEncode(rs("skinshowlog")) else Response.Write("")%></textarea>
        </strong></td>
    </tr>
    <tr>
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存修改 " >
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
set rs=nothing
end sub

sub addskin()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">添加系统模板</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
  <tr class="tdbg">
    <td height="30"><div align="center">
	<a href="admin_sysskin.asp?action=showskin"><strong>返回管理菜单</strong></a>
	<a href="admin_skin_help.asp" target="_blank"><strong>模板标记帮助</strong></a>
	</div></td>
  </tr>
</table>

<form method="POST" action="admin_sysskin.asp" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>模板参数</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">模板名称：
        <input name="sysskinname" type="text" id="sysskinname">
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <strong>主模板：</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"></textarea>
        <br>
        <br>
        <strong>副模板： <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"></textarea>
        </strong></td>
    </tr>
    <tr>
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveaddskin">
          <input name="cmdadd" type="submit" id="cmdadd" value=" 添加 " >
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
end Sub
Set oblog = Nothing
%>