<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_skin_sys")=False Then Response.Write "��Ȩ����":Response.End
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
		<li class="main_top_left left">ϵͳģ�����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form name="form2" method="post" action="m_sysskin.asp?action=savedefault">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
  <tr class="topbg">
    <td height="25" align="center" width="25"><strong>ID</strong></td>
    <td align="center" width="120"><strong>����</strong></td>
    <td align="center" width="100"><strong>����</strong></td>
	<td align="center" width="60"><strong>Ĭ��ģ��</strong></td>
    <td align="center"><strong>ģ�����</strong></td>
  </tr>

 <%
while not rs.eof
%>
          <tr class="tdbg">
          <td>
              <div align="center"><%= rs("id") %>&nbsp;</div></td>
          <td> <div align="center"><div align="center"><%If rs("isdefault")=1 Then Response.Write "<font color=red>" & rs("sysskinname") & "</font>" Else Response.Write rs("sysskinname") End If %></div></td>
          <td><div align="center"><%= rs("skinauthor") %></div></td>
            <td>              <div align="center">
                <input name="radiobutton" type="radio" class="tdbg" value='<%=rs("id")%>' <%if rs("isdefault")=1 then Response.Write "checked" %>>
            </div></td>
            <td>
                <div align="center">
				<a href="../admin_edit.asp?action=modiskin&skintype=sys&t=0&editm=1&skinorder=0&id=<%=rs("id")%>"  target="_blank">�޸���ģ��</a>
                        ��<a href="../admin_edit.asp?action=modiskin&skintype=sys&t=0&editm=1&skinorder=1&id=<%=rs("id")%>" target="_blank">�޸ĸ�ģ��</a>
				<a href="m_sysskin.asp?action=modiskin&id=<%=rs("id")%>">�޸�ģ��(�ı���ʽ)</a>
          ��<a href="m_sysskin.asp?action=delconfig&id=<%=rs("id")%>" onclick=return(confirm("ȷ��Ҫɾ�����ģ����"))>ɾ��ģ��</a></div></td>
        </tr>
      <%
rs.movenext
wend
%>
    <tr>
    <td height="40" colspan="5" align="center" class="tdbg"> <div align="center">
          <input type="submit" name="Submit" value="��������">
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
	if Trim(Request("sysskinname"))="" then oblog.sys_err("ģ��������Ϊ��"):Response.End()
	if Trim(Request("skinmain"))="" then oblog.sys_err("�����治��Ϊ��"):Response.End()
	if Trim(Request("skinshowlog"))="" then oblog.sys_err("�����治��Ϊ��"):Response.End()
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
	WriteSysLog "�������޸�ϵͳģ��������ı���ʽ����Ŀ��ģ��ID��"&Request.QueryString("id")&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "�޸ĳɹ�",""
end sub
sub savedefault()
	dim isdefaultID
	isdefaultID=CLng(Trim(Request("radiobutton")))
	oblog.execute("update oblog_sysskin set isdefault=0")
	oblog.execute("update oblog_sysskin set isdefault=1 where id="&isdefaultID)
	WriteSysLog "�������趨ϵͳĬ��ģ�������Ŀ��ģ��ID��"&isdefaultID&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""�޸ĳɹ���"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub delconfig()
	oblog.execute("delete from oblog_sysskin where id="&CLng(Request.QueryString("id")))
	WriteSysLog "������ɾ��ϵͳģ�������Ŀ��ģ��ID��"&Request.QueryString("id")&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Redirect "m_sysskin.asp?action=showskin"
end sub

sub saveaddskin()
	dim rs,sql
	set rs=Server.CreateObject("adodb.recordset")
	if Trim(Request("sysskinname"))="" then oblog.sys_err("ģ��������Ϊ��"):Response.End()
	if Trim(Request("skinmain"))="" then oblog.sys_err("��ģ�治��Ϊ��"):Response.End()
	if Trim(Request("skinshowlog"))="" then oblog.sys_err("��ģ�治��Ϊ��"):Response.End()
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
	WriteSysLog "���������ϵͳģ�����",oblog.NowUrl&"?"&Request.QueryString
	Response.Redirect "m_sysskin.asp?action=showskin"
end sub

sub modiskin()
	dim rs
	set rs=oblog.execute("select * from oblog_sysskin where id="&CLng(Request.QueryString("id")))
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
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
    <tr class="topbg">

    <td height="22" colspan=2 align=center><strong>�޸�ϵͳģ��</strong></td>
    </tr>
    <tr class="tdbg">

    <td width="253" height="30"><strong>�����޸ĵ�ģ���ǣ�<%=rs("sysskinname")%></strong></td>

    <td width="516" height="30"><a href="m_sysskin.asp?action=modiskin&id=<%=rs("id")%>">�޸�ģ��</a>����<a href="m_sysskin.asp?action=showskin">���ع���˵�</a>
����<a href="m_sysskin.asp?action=showskin">���ع���˵�</a>
      <a href="m_skin_help.asp" target="_blank"><strong>ģ���ǰ���</strong></a>
	 </td>
    </tr>
</table>

<form method="POST" action="m_sysskin.asp?id=<%=CLng(Request.QueryString("id"))%>" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="769" height="22" class="topbg"><strong>ģ�����</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">ģ�����ƣ�
        <input name="sysskinname" type="text" id="sysskinname" value=<%=rs("sysskinname")%>>
        �������ߣ�
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <strong>��ģ�棺</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"><%if rs("skinmain")<>"" then Response.Write Server.HtmlEncode(rs("skinmain")) else Response.Write("")%></textarea>
        <br>
        <br>
        <strong>��ģ�棺 <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"><%if rs("skinshowlog")<>"" then Response.Write Server.HtmlEncode(rs("skinshowlog")) else Response.Write("")%></textarea>
        </strong></td>
    </tr>
    <tr>
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" �����޸� " >
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
		<li class="main_top_left left">���ϵͳģ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
  <tr class="tdbg">
    <td height="30"><div align="center">
	<a href="m_sysskin.asp?action=showskin"><strong>���ع���˵�</strong></a>
	<a href="m_skin_help.asp" target="_blank"><strong>ģ���ǰ���</strong></a>
	</div></td>
  </tr>
</table>

<form method="POST" action="m_sysskin.asp" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="769" height="22" class="topbg"><strong>ģ�����</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">ģ�����ƣ�
        <input name="sysskinname" type="text" id="sysskinname">
        �������ߣ�
        <input name="skinauthor" type="text" id="skinauthor"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <strong>��ģ�棺</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"></textarea>
        <br>
        <br>
        <strong>��ģ�棺 <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"></textarea>
        </strong></td>
    </tr>
    <tr>
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveaddskin">
          <input name="cmdadd" type="submit" id="cmdadd" value=" ��� " >
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