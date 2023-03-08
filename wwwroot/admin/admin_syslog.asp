<!--#include file="inc/inc_sys.asp"-->
<%
dim rs, sql,action,FoundErr
dim id,cmd,Keyword,strField,date1,date2
keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
strField=Trim(Request("Field"))
cmd=Trim(Request("cmd"))
Action=Trim(Request("Action"))
id=Trim(Request("id"))
date1=DeDateCode(Request("date1"))
date2=DeDateCode(Request("date2"))
if cmd="" then
	cmd=0
else
	cmd=CLng(cmd)
end if
G_P_FileName="admin_syslog.asp?cmd=" & cmd
If Keyword <>"" Then
	G_P_FileName = G_P_FileName & "&keyword="&keyword&"&Field="&strField
Else
	If date1 <> "" Or date2<>"" Then
		G_P_FileName= G_P_FileName &"&date1="&Request("date1")&"&date2="&Request("date2")
	End If
End if
if Request("page")<>"" then
    G_P_This=cint(Request("page"))
else
	G_P_This=1
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>系统日志管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<SCRIPT language=javascript>
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
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}

function CheckSelect(form)
{
  var j;
  j=0
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
    	if (e.checked ) {
    		j=j+1;
    		break;
    	}
    }
    if(j>0) {
    	return true;
    }
    else{
    	alert("必须选择相关数据才能进行操作")
    	return false;
    }

}
</SCRIPT>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">系统日志管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="form1" method="post" action="admin_syslog.asp">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查询：</strong></td>
      <td width="687" height="30">
      	<select size=1 name="cmd">
          <option value=>请选择查询条件</option>
		  <option value="1">错误登录日志</option>
		  <option value="2">系统管理员操作日志</option>
          <option value="3">用户敏感操作日志</option>
          <option value="4">内容管理员操作日志</option>
        </select>
        <input type="submit" value=" 查 询 ">
      </td>
    </tr>
</form>
  <form name="form3" method="post" action="admin_syslog.asp">
  <tr class="tdbg">
      <td width="120"><strong>高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">
      <option value="username" selected>管理员名称(用户ID)</option>
      <option value="ip" >登录IP</option>
	  <option value="userid" >可疑用户ID</option>
      </select>
	  <input name="cmd" type="hidden" id="cmd" value="10">
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
        若为空，则查询所有 &nbsp;&nbsp;|&nbsp;&nbsp;<a href="#" onclick="if(confirm('确认要清空所有登录日志吗?')==true) document.location.href='admin_syslog.asp?action=clearlog';">清空所有日志</td>
  </tr>
</form>
  <form name="form2" method="post" action="admin_syslog.asp">
  <tr class="tdbg">
      <td width="120"><strong>按时间区段查询：</strong></td>
    <td>
    	开始时间：<input type="text" name="date1" size=12 maxlength=8>
    	结束时间：<input type="text" name="date2" size=12 maxlength=8>

      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="cmd" type="hidden" id="cmd" value="11">
      <br/>
        时间格式：YYYYMMDD，如2006年6月6日，则输入20060606,其他格式均不支持</td>
  </tr>
</form>
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
	Case "del"
		Call Dellogs
	Case "clearlog"
		Call Delalllogs
	Case else
		call main()
end select
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	G_P_Guide="<table width='100%'><tr><td align='left'>您现在的位置：<a href='admin_syslog.asp'>日志管理</a>&nbsp;&gt;&gt;&nbsp;"
	select case cmd
		case 0
			sql="select top 500 * from oblog_syslog order by id desc"
			G_P_Guide=G_P_Guide & "最后500条操作日志"
		case 1
			sql="select top 500 * from oblog_syslog  WHERE itype = 0 order by id desc"
			G_P_Guide=G_P_Guide & "最后500条错误登录日志"
		case 2
			sql="select top 500 * from oblog_syslog WHERE itype = 1 order by id desc"
			G_P_Guide=G_P_Guide & "最后500条系统管理员操作日志"
		case 3
			sql="select top 500 * from oblog_syslog WHERE itype = 2 order by id desc"
			G_P_Guide=G_P_Guide & "最后500条用户敏感操作日志"
		case 4
			sql="select top 500 * from oblog_syslog WHERE itype = 3 order by id desc"
			G_P_Guide=G_P_Guide & "最后500条内容管理员操作日志"
		case 10
			if Keyword="" Then
				If strField = "userid" Then
					sql="select top 500 * from oblog_syslog WHERE itype=2 order by id desc"
					G_P_Guide=G_P_Guide & "最后500个用户发布的可疑日志"
				Else
					sql="select top 500 * from oblog_syslog order by id desc"
					G_P_Guide=G_P_Guide & "最后500条操作日志"
				End if
			else
				select case strField
					case "userid"
						If Not IsNumeric(Keyword) Then
							Oblog.ShowMsg "用户ID必须为整数",""
						End if
						Keyword=CLng(Keyword)
						sql="select * from oblog_syslog where username = '" & Keyword & "'"
						G_P_Guide=G_P_Guide & "可疑用户ID为<font color=red> " & Keyword & " </font>的操作日志"
					case "username"
						sql="select * from oblog_syslog where username like '%" & Keyword & "'"
						G_P_Guide=G_P_Guide & "管理员名称中包含<font color=red> " & Keyword & " </font>的操作日志"
					case "ip"
						sql="select top 500 * from oblog_syslog where addip like '%" & Keyword & "%' order by id  desc"
						G_P_Guide=G_P_Guide & "IP中包含<font color=red> " & Keyword & " </font>的操作日志"

				end select
			end if
		Case 11
			If date1<>"" And date2<>"" Then
				Sql="select * From oblog_syslog Where addtime>=" & G_Sql_d_Char & date1 & G_Sql_d_Char & " And addtime<=" & G_Sql_d_Char&  date2 & G_Sql_d_Char
				G_P_Guide=G_P_Guide & "自 " & date1 & " 至 " & date2 & " 期间的操作日志"
			End If
		case else
			sql="select top 500 * from oblog_syslog order by id desc"
			G_P_Guide=G_P_Guide & "最后500次操作日志"
	end select
	G_P_Guide=G_P_Guide & "</td><td align='right'>"
	If SQL = "" Then
		oblog.ShowMsg "查询参数不正确",""
	End If
	'if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'OB_DEBUG SQL,1
	rs.Open sql,Conn,1,1
	'Response.Write G_P_Guide
	Response.Write "<br/>"
  	Call oblog.MakePagebar(rs,"条操作日志")
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i
    i=0
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">系统日志管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="myform" method="Post" action="admin_syslog.asp?action=del" onsubmit="return confirm('确定要执行选定的操作吗？');">
<style type="text/css">
<!--
td {padding:3px 0!important;}
-->
</style>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
<%do while not rs.EOF %>
  <tr>
    <td align="center" width="30" style="background:#ccc;border-bottom:1px #666 dotted;">
<input type="checkbox" name="id" value="<%=rs("id")%>">
	</td>
<%If rs("itype")=2 Then%>
    <td style="background:#ededed;border-bottom:1px #888 dotted;">
　<span style="color:#f00;">可疑用户：<%=rs("username")%></span>
    </td>
    <td style="background:#ededed;border-bottom:1px #888 dotted;" width="160">
发布时间：<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;color:#666;"><%=rs("addtime")%></span>
	</td>
<%Else%>
    <td style="background:#ededed;border-bottom:1px #888 dotted;">
　管理员名称：<span style="color:#0D4D89;"><%=rs("username")%></span>
    </td>
    <td style="background:#ededed;border-bottom:1px #888 dotted;" width="160">
登录时间：<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;color:#666;"><%=rs("addtime")%></span>
	</td>
<%End if%>
    <td style="background:#ededed;border-bottom:1px #888 dotted;" width="140">
登录IP：<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;color:#666;"><%=rs("addip")%></span>
</td>
  </tr>
    <tr>
	 <td align="center" valign="top"></td>
    <td colspan="3" height="8">操作连接：<%=OB_IIF(rs("QueryStrings"),"省略")%></td>
  </tr>
  <tr>
    <td align="center" valign="top"></td>
    <td colspan="3" valign="top" style="word-break:break-all;color:#555;"><%If rs("itype")=0 Then%><span style="color:#f60;"><%=rs("desc")%></span><%else%><%=rs("desc")%><%End if%></td>
  </tr>
          <%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
%>
  <tr>
    <td colspan="4" valign="top" align="center"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox"> 选中全部进行删除 <input type="submit" name="Submit" value=" 执 行 " onclick="return CheckSelect(this.form)"></td>
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
End sub

Sub Dellogs()
	Dim id
	id=Request("id")
	id=FilterIds(id)
	If Id="" Then exit Sub
	oblog.execute("delete from oblog_syslog where DATEDIFF("&G_Sql_H&",addtime,"&G_Sql_Now&") > 72 AND id in("&id & ")")
	EventLog "删除了部分操作日志，最近72小时的操作日志将被保留！",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "删除成功，最近72小时的操作日志将被保留！",""
End Sub

Sub Delalllogs()
	oblog.execute("Delete From oblog_syslog WHERE DATEDIFF("&G_Sql_H&",addtime,"&G_Sql_Now&") > 72 ")
	EventLog "对日志进行了清空，最近72小时的操作日志将被保留！",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "系统操作日志清除成功，最近72小时的操作日志将被保留！","admin_syslog.asp"
End Sub
Set oblog = Nothing
%>