<!--#include file="inc/inc_sys.asp"-->
<%
Dim Action
Action=Trim(Request("Action"))
if Action="saveconfig" then
	call saveconfig()
else
	call showconfig()
end if

sub showconfig()
dim rs,blackIps

'������
set rs=oblog.execute("select * from oblog_config Where id=10")
If Not rs.Eof Then
	blackIps=Ob_IIF(rs(1),"")
End If
Set rs=Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>blog��ӳ��������</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">��ӳ�������</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="25" class="tdbg">����������û���ӳ��������࣬�û����Ķ���־��ʱ����������漰����Ƿ�����ɫ�����ε����ݿ��Լ�ʱ��������վ�����Ա㼰ʱ����������ɲ���Ҫ����ʧ��
      <br/>
	  <font color=red>����ѡ���Իس��ֿ�</font>
      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <textarea name="ips1" cols="35" rows="15" id="lockip">
<%=blackIps%></textarea>
      </td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">

    <tr>
      <td height="40" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="saveconfig"> <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " >
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
</form>
</body>
</html>
<%
set rs=nothing
end sub

sub saveconfig()
'	OB_DEBUG Request.QueryString,1
	If Request.QueryString <> "" Then Exit Sub
	dim rs,sql
	if not IsObject(conn) then link_database
	set rs=Server.CreateObject("adodb.recordset")
	rs.Open "select * From  oblog_config Where id=10",conn,1,3
	If rs.Eof Then
		rs.AddNew
		rs("id")=10
	End If
	rs("ob_value")=oblog.FilterEmpty(Request("ips1"))
	rs.Update
	rs.Close
	Set rs=Nothing
	oblog.ReloadCache
	EventLog "�����˷�ӳ����������",""
	Response.Redirect "admin_report.asp"
end Sub
Set oblog = Nothing
%>