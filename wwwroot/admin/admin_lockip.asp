<!--#include file="inc/inc_sys.asp"-->
<%
Dim action
Action=Trim(Request("Action"))
if Action="saveconfig" then
	call saveconfig()
else
	call showconfig()
end if

sub showconfig()
dim rs,whiteIps,blackIps
'������
set rs=oblog.execute("select id,ob_value from oblog_config Where id=4")
If Not rs.Eof Then
	whiteIps=Ob_IIF(rs(1),"")
End If
'������
set rs=oblog.execute("select * from oblog_config Where id=5")
If Not rs.Eof Then
	blackIps=Ob_IIF(rs(1),"")
End If
Set rs=Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>blog��־��������</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">��վ����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="admin_lockip.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"> <strong>IP���ƹ���(������),����Ա���ɲ���</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">��������Ӷ������IP��ÿ��IP�ûس��ָ�������IP����д��ʽ��202.152.12.1��������202.152.12.1���IP�ķ��ʣ���202.152.12.*����������202.152.12��ͷ��IP���ʣ�ͬ��*.*.*.*������������IP�ķ��ʡ�����Ӷ��IP��ʱ����ע�����һ��IP�ĺ��治Ҫ�ӻس�
      <br/>
      <font color=red><strong>�˴����Զ������ظ��������IP����ʾ��Ϻ�����ذ����棬��������Ч��</strong></font>
      </td>
</tr>
    <tr>
      <td height="25" class="tdbg"> <textarea name="ips1" cols="35" rows="15" id="lockip">
<%
'�ڴ˴�����һ���������ˣ���ֹ�ظ�IP����
Dim ip0,ip1,ip2,i,j,z
j=0
ip0=blackIps
ip2="||"
ip1=Split(ip0,VBcrlf)
For i=0 To Ubound(ip1)
	If Len(Trim(ip1(i)))>=7 And Len(Trim(ip1(i)))<=15 Then
		If Instr(ip2,"||" & ip1(i) & "||")<=0 Then
			ip2=ip2  & ip1(i) & "||"
			'Response.Write ip1(i) & "<br/>"
			j=j+1
		ENd If
	ENd If
Next
'ȥ��ͷ��β
If Len(Ip2)>=7 Then ip2=Mid(Ip2,3,Len(ip2)-4)
Response.Write Replace(ip2,"||",vbcrlf)
%>
</textarea>
      </td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"> <strong>IP������,��ϵͳ����Ա���Բ���</strong></td>
    </tr>
    <tr>
      <td height="22" class="tdbg">��Ϊ���ݹ���Ա�������ú�����,Ϊ�˷�ֹ���ݹ���Ա����������Ա��IP������������³�������ԱIP�޷�������ز���,����Ա���Խ��Լ���IP����Ϊ�����������ܺ�����Լ��</td>
    </tr>
    <tr>
      <td height="25" class="tdbg">������Ӷ��IP��ÿ��IP�ûس��ָ���������IP�����ں�������������ֻ����ϵͳ����Ա����</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <textarea name="ips2" cols="35" rows="10" id="lockip">
<%=whiteIps%></textarea>
      </td>
    </tr>
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
	If Request.QueryString <> "" Then Exit Sub
	dim rs,sql
	if not IsObject(conn) then link_database
	set rs=Server.CreateObject("adodb.recordset")
	rs.Open "select * From  oblog_config Where id=4",conn,1,3
	If rs.Eof Then
		rs.AddNew
		rs("id")=4
	End If
	rs("ob_value")=oblog.FilterEmpty(Request("ips2"))
	rs.Update
	rs.Close

	rs.Open "select * From  oblog_config Where id=5",conn,1,3
	If rs.Eof Then
		rs.AddNew
		rs("id")=5
	End If
	rs("ob_value")=oblog.FilterEmpty(Request("ips1"))
	rs.Update
	rs.Close
	Set rs=Nothing
	oblog.ReloadCache
	EventLog "����������IP�������",""
	Response.Redirect "admin_lockip.asp"
end Sub
Set oblog = Nothing
%>