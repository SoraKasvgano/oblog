<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_IP")=False Then Response.Write "无权操作":Response.End
Action=Trim(Request("Action"))
if Action="saveconfig" then
	call saveconfig()
else
	call showconfig()
end if

sub showconfig()
dim rs,blackIps

'黑名单
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
<title>blog日志过滤设置</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">IP限制管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="m_lockip.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="25" class="tdbg">您可以添加多个限制IP，每个IP用回车分隔，限制IP的书写方式如202.152.12.1就限制了202.152.12.1这个IP的访问，如202.152.12.*就限制了以202.152.12开头的IP访问，同理*.*.*.*则限制了所有IP的访问。在添加多个IP的时候，请注意最后一个IP的后面不要加回车
      <br/>
	<font color=red><strong>此处会自动清理重复、错误的IP，显示完毕后请务必按保存，提升阻拦效率</strong></font>
      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <textarea name="ips1" cols="35" rows="15" id="lockip">
<%'在此处进行一次主动过滤，防止重复IP出现
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
'去除头和尾
If Len(Ip2)>=7 Then ip2=Mid(Ip2,3,Len(ip2)-4)
Response.Write Replace(ip2,"||",vbcrlf)
%></textarea>
      </td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">

    <tr>
      <td height="40" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="saveconfig"> <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " >
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
	WriteSysLog "进行了限制IP管理操作",""
	oblog.ShowMsg "操作成功","m_lockip.asp"
end Sub
Set oblog = Nothing
%>