<!--#include file="inc/inc_sys.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>�޸Ĺ���Ա����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function check()
{
  if(document.form1.Pwd.value=="")
    {
      alert("ԭ���벻��Ϊ�գ�");
	  document.form1.Pwd.focus();
      return false;
    }
  if(document.form1.Password.value=="")
    {
      alert("���벻��Ϊ�գ�");
	  document.form1.Password.focus();
      return false;
    }

  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("��ʼ������ȷ�����벻ͬ��");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();
      return false;
    }
}
</script>
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<body class="bgcolor">
<%
dim rs,sql
Action=Trim(Request("Action"))
If Session("adminname")<>"" Then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>ϵͳ����Ա�����ڴ˽������޸����룡</li>"
	call WriteErrMsg()
	Response.End
End If
sql="select * from oblog_Admin where UserName='" & ProtectSql(Session("m_name")) & "' And roleid>0"
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open sql,conn,1,3
if rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>����Ա�������û������ڴ˽���������룡</li>"
else
	if Action="Modify" then
		call ModifyPwd()
	else
		call main()
	end if
end if
rs.close
set rs=nothing
if FoundErr=True then
	call WriteErrMsg()
end if
Set oblog = Nothing
sub ModifyPwd()
	dim password,PwdConfirm,pwd
	password=Trim(Request("Password"))
	PwdConfirm=Trim(Request("PwdConfirm"))
	pwd=Trim(Request("pwd"))
	If MD5(pwd) <> rs("password") Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ԭ�����������</li>"
		exit sub
	End if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����벻��Ϊ�գ�</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ȷ�������������������ͬ��</li>"
		exit sub
	end if
	if Password<>"" then
		rs("password")=md5(password)
	end if
   	rs.update
	WriteSysLog "�����������޸Ĳ���",oblog.NowUrl&"?"&Request.QueryString
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""�����޸ĳɹ��������µ�¼��"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub main()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�޸Ĺ���Ա����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">

<form method="post" action="m_pwd.asp" name="form1" onsubmit="javascript:return check();">
  <table width="300" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border" >
    <tr class="title">
      <td height="25" colspan="2" class="topbg"><strong>�޸Ĺ���Ա����</strong></td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg">�� �� ����</td>
      <td class="tdbg"><%=rs("UserName")%></td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg">ԭ �� �룺</td>
      <td class="tdbg"><input type="password" name="Pwd"> </td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg">�� �� �룺</td>
      <td class="tdbg"><input type="password" name="Password"> </td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg">ȷ�����룺</td>
      <td class="tdbg"><input type="password" name="PwdConfirm"> </td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="Modify">
        <input  type="submit" name="Submit" value=" ȷ �� " style="cursor:hand;">
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="reset()" style="cursor:hand;"></td>
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
end sub

sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbcrlf
	strErr=strErr & "<head>" & vbcrlf
	strErr=strErr & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbcrlf
	strErr=strErr & "<title>oBlog--��̨����</title>" & vbcrlf
	strErr=strErr & "<link rel=""stylesheet"" href=""images/style.css"" type=""text/css"" />" & vbcrlf
	strErr=strErr & "</head>" & vbcrlf
	strErr=strErr & "<body>" & vbcrlf
	strErr=strErr & "<div id=""main_body"">" & vbcrlf
	strErr=strErr & "	<ul class=""main_top"">" & vbcrlf
	strErr=strErr & "		<li class=""main_top_left left"">������Ϣ</li>" & vbcrlf
	strErr=strErr & "		<li class=""main_top_right right""> </li>" & vbcrlf
	strErr=strErr & "	</ul>" & vbcrlf
	strErr=strErr & "	<div class=""main_content_rightbg"">" & vbcrlf
	strErr=strErr & "		<div class=""main_content_leftbg"">" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "		</div>" & vbcrlf
	strErr=strErr & "	</div>" & vbcrlf
	strErr=strErr & "	<ul class=""main_end"">" & vbcrlf
	strErr=strErr & "		<li class=""main_end_left left""></li>" & vbcrlf
	strErr=strErr & "		<li class=""main_end_right right""></li>" & vbcrlf
	strErr=strErr & "	</ul>" & vbcrlf
	strErr=strErr & "</div>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	Response.write strErr
end sub
%>