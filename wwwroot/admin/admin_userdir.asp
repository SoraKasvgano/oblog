<!--#include file="inc/inc_sys.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>վ������</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<%
dim action
action=Request.QueryString("action")
select case action
	case "adduserdir"
		call adduserdir()
	case "del"
		call deluserdir()
	case "setdefault"
		call setdefault()
	case else
		call main
end select

sub adduserdir()
	dim userdir,rs,oFSO
	userdir=oblog.filt_badstr(Trim(Request.Form("userdir")))
	If userdir="" Then
		Response.Write("<script language=javascript>alert('Ŀ¼����Ϊ�գ�');window.location.replace('admin_userdir.asp')</script>")
		Response.End()
	End If
	If oblog.chkdomain(userdir) = False Then
		oblog.ShowMsg "�û�Ŀ¼������ʹ�������ַ�",""
	End If
	Dim arrayDir,i
	arrayDir = Oblog.SysDir
	For i = 0 To UBound(arrayDir)
		if LCase(userdir) = arrayDir(i) Then
			oblog.ShowMsg "����ѡ��ϵͳĿ¼��Ϊ�û�Ŀ¼",""
		End If
	Next
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open "select * From oblog_userdir Where userdir='" &userdir & "'" ,conn,1,3
	If Not rs.Eof Then
		rs.Close
		Set rs=Nothing
		Response.Redirect "admin_userdir.asp"
	End If
	rs.Close
	Set rs=Nothing
	oblog.execute("insert into [oblog_userdir] (userdir,is_default) values ('"&userdir&"',0)")
	On Error Resume Next
	'�ж�Ŀ¼�Ƿ���ڣ�������������Զ�����
	Set oFso=Server.CreateObject(oblog.CacheCompont(1))
	If oFso.FolderExists(Server.Mappath(blogdir & userdir)) =false Then oFso.CreateFolder(Server.Mappath(blogdir & userdir))
	Set oFso=Nothing
	If Err Then
		Err.Clear
		oblog.ShowMsg "�û�Ŀ¼����ʧ�ܣ����ֹ�����",""
	End if
	EventLog "���д����û�Ŀ¼�Ĳ�����Ŀ���û�Ŀ¼Ϊ��"&userdir&"",oblog.NowUrl&"?"&Request.QueryString
    Response.Redirect "admin_userdir.asp"
end sub

sub deluserdir()
	dim id
	id=CLng(Request.QueryString("id"))
	oblog.execute("delete  from [oblog_userdir] where id="&id)
	EventLog "����ɾ���û�Ŀ¼�Ĳ�����Ŀ���û�Ŀ¼IDΪ��"&id&"",oblog.NowUrl&"?"&Request.QueryString
    Response.Redirect "admin_userdir.asp"
end sub

sub setdefault()
	dim id,rs
	id=CLng(Request.QueryString("id"))
	oblog.execute("update [oblog_userdir] set is_default=0")
	oblog.execute("update [oblog_userdir] set is_default=1 where id="&id)
	set rs=oblog.execute("select userdir from oblog_userdir where is_default=1")
	oblog.execute("update oblog_setup set user_dir='"&rs(0)&"' where id=1")
	set rs=nothing
	oblog.ReloadSetup
	EventLog "�����趨Ĭ���û�Ŀ¼�Ĳ�����Ŀ���û�Ŀ¼IDΪ��"&id&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Redirect "admin_userdir.asp"
end sub



sub main()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�û�Ŀ¼����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">




<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr valign="middle">
    <td height="21" colspan="3" class="topbg"> <strong>����û�Ŀ¼</strong></td>
  </tr>
   <tr class="tdbg"><form name="form1" method="post" action="admin_userdir.asp?action=adduserdir">
          <td height="20" colspan="3"><div align="center">Ŀ¼����
          <input name="userdir" type="text" id="userdir" maxlength="20">
          <input type="submit" name="Submit" value=" ��� ">
          </div></td></form>
  </tr>
</table>

<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="topbg" height="25">
    <td width="38" > <div align="center">ID</div></td>
    <td width="109"> <div align="center">Ŀ¼��</div></td>
    <td width="72"> <div align="center">�û���</div></td>
    <td width="110"><div align="center">��ǰʹ��Ŀ¼</div></td>
    <td width="242"> <div align="center">�������</div></td>
  </tr>
  <%
dim rs,rstmp
set rs=oblog.execute("select * from oblog_userdir order by id")
while not rs.eof
%>
  <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
    <td > <div align="center"><%=rs("id")%></div></td>
    <td><div align="center"><%=rs("userdir")%></div></td>
	<td><div align="center">
<%
set rstmp=oblog.execute("select count(userid) from oblog_user where user_dir='"&rs("userdir")&"'")
Response.Write(rstmp(0))
%></div></td>

    <td> <div align="center">
    <%if rs("is_default")=1 then Response.Write "<font color=red>��</font>" else Response.Write("��")%>
      </div></td>
    <td><div align="center"><a href="admin_userdir.asp?action=setdefault&id=<%=rs("id")%>" <%="onClick='return confirm(""ȷ�ϴ�Ŀ¼Ϊ�û�Ĭ��Ŀ¼��"");'"%>>����ΪĬ��</a>
        | <a href="admin_userdir.asp?action=del&id=<%=rs("id")%>" <%="onClick='return confirm(""ɾ���󣬴˳�Ա������־�����������blog����ʾ,ȷ��Ҫɾ����"");'"%>>ɾ��</a></div></td>
  </tr>
  <%
rs.movenext
wend
%>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%
end sub
Set oblog=Nothing
%>
</body>