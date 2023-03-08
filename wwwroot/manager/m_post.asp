<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_group_blog")=False Then Response.Write "无权操作":Response.End
dim rs, sql
dim id,cmd,Keyword,sField,sDate1,sDate2
keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
sField=Trim(Request("Field"))
cmd=Trim(Request("cmd"))
Action=Trim(Request("Action"))
id=Trim(Request("id"))
sDate1=Request("date1")
sDate2=Request("date2")
If sDate1<>"" Then sDate1=Int(sDate1)
If sDate2<>"" Then sDate2=Int(sDate2)
if cmd="" then
	cmd=0
else
	cmd=CLng(cmd)
end if
G_P_FileName="m_post.asp?cmd=" & cmd & "&field=" & sField & "&keyword=" & keyword & "&date1=" & sDate1 & "&date2=" &sDate2
if Request("page")<>"" then
    G_P_This=cint(Request("page"))
else
	G_P_This=1
end if

%>
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
</SCRIPT>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">群 组 内 容 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_post.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查找：</strong></td>
      <td width="687" height="30"><a href="m_post.asp">所有内容</a>|&nbsp;&nbsp;<a href="m_post.asp?cmd=1">主题列表</a>|&nbsp;&nbsp;<a href="m_post.asp?cmd=3">精华列表</a>|&nbsp;&nbsp;<a href="m_post.asp?cmd=2">回复列表</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="m_post.asp">
  <tr class="tdbg">
      <td width="120"><strong>高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">
	      <option value="author" selected>用户名称</option>
	      <option value="userid" >用户ID</option>
	      <option value="ip" selected>发表IP</option>
	      <option value="title" >标题内容</option>
	      <option value="content" >正文内容</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="cmd" type="hidden" id="cmd" value="10">
        若为空，则查询所有</td>
  </tr>
</form>
  <form name="form3" method="post" action="m_post.asp">
  <tr class="tdbg">
      <td width="120"><strong>按时间区段查询：</strong></td>
    <td>
    	开始时间：<input type="text" name="date1" size=12 maxlength=10>
    	结束时间：<input type="text" name="date2" size=12 maxlength=10>

      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="cmd" type="hidden" id="cmd" value="11">
      <br/>
        时间格式：YYYYMMDDHHmm，如2006年6月6日9点，则输入2006060609,其他格式均不支持</td>
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
If id<>"" Then
	If Instr(id,",") Then
		id=FilterIds(id)
	Else
		id=Int(Id)
	End If
End If
select Case LCase(action)
	Case "del"
		oblog.execute("delete from oblog_teampost where postid In ("&id & ")")
		oblog.execute("delete from oblog_teampost where parentid In ("&id & ")")
		WriteSysLog "进行了删除"&oblog.CacheConfig(69)&"主题操作，目标"&oblog.CacheConfig(69)&"ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "删除成功！",""
	Case Else
		call main()
end select
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	Dim sQryFields
	sQryFields="top 500 topic,content,postid,logid,teamid,userid,addtime,isbest,istop,idepth,author,addip,parentid"
	select case cmd
		case 0
			sql="select " & sQryFields &" from oblog_teampost order by postid desc"
			sGuide=sGuide & "最后500篇内容"
		Case 1
			sql="select " & sQryFields &" from oblog_teampost Where idepth=0 order by postid desc"
		Case 2
			sql="select " & sQryFields &" from oblog_teampost Where idepth>0 order by postid desc"
		Case 3
			sql="select " & sQryFields &" from oblog_teampost Where idepth=0 And isbest=1 order by postid desc"
		case 10
			if Keyword="" then
				sql="select " & sQryFields &" from oblog_teampost order by postid desc"
				sGuide=sGuide & "所有日志"
			else
				select case sField
				case "userid"
					if Not IsNumeric(Keyword) then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID必须是整数！</li>"
					else
						sql="select " & sQryFields & " from oblog_teampost where userid =" & CLng(Keyword)
						sGuide=sGuide & "作者ID等于<font color=red> " & CLng(Keyword) & " </font>的日志"
					end if
				case "author"
					sql="select " & sQryFields & " from oblog_teampost where author like '%" & Keyword & "%' order by postid  desc"
					sGuide=sGuide & "作者名称中含有“ <font color=red>" & Keyword & "</font> ”的日志"
				case "ip"
					sql="select " & sQryFields & " from oblog_teampost where addip like '%" & Keyword & "%' order by postid  desc"
					sGuide=sGuide & "发布日志时的IP中含有“ <font color=red>" & Keyword & "</font> ”的日志"
				case "title"
					sql="select " & sQryFields & " from oblog_teampost where topic like '%" & Keyword & "%' order by postid  desc"
					sGuide=sGuide & "标题中含有“ <font color=red>" & Keyword & "</font> ”的日志"
				case "content"
					sql="select " & sQryFields & " from oblog_teampost where content like '%" & Keyword & "%' order by postid  desc"
					sGuide=sGuide & "内容中含有“ <font color=red>" & Keyword & "</font> ”的日志"
				end select
			end if
		Case 11
			sDate1=DeDateCode(sDate1)
			sDate2=DeDateCode(sDate2)
			If sDate1<>"" And sDate2<>"" Then
				sql="select " & sQryFields & " from oblog_teampost where addtime>=" & G_Sql_d_Char & sDate1 & G_Sql_d_Char & " And  addtime<=" & G_Sql_d_Char & sDate2 & G_Sql_d_Char &  " order by postid  desc"
				sGuide=sGuide & "实际发布时间在 <font color=red>" & sDate1 & "</font> 至 <font color=red>" & sDate2 & "</font> 的内容"
			End If
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>错误的参数！</li>"
	end select
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'Response.Write sql
	rs.Open sql,Conn,1,1
  Call oblog.MakePagebar(rs,"篇")
end sub

sub showContent()
   	dim i
    i=0
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left"><%=sGuide%></li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<style type="text/css">
<!--
td {padding:3px 0!important;}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" style="word-wrap: break-word;word-break:break-all;">
  <form name="myform" method="Post" action="m_post.asp" onsubmit="return confirm('确定要执行选定的操作吗？');">
          <%do while not rs.EOF %>
  <tr>
    <td align="center" style="background:#B3D1EA;border-bottom:1px #000 dotted;" width="30">
    	<input name='id' type='checkbox' onclick="unselectall()" id="id" value='<%=cstr(rs("postid"))%>'>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;">
    <span style="margin:0 0 0 10px;">
	<%If rs("idepth")=0 Then Response.Write "<a href=""m_post.asp?cmd=1"" style=""color:#217DBD;font-weight:600;"">[主题]</a> " Else Response.Write "<a href=""m_post.asp?cmd=2"" style=""color:#99F;font-weight:600;"">[回复]</a> " End if%>
    <%If rs("isbest")=1 Then Response.Write "<a href=""m_post.asp?cmd=3"" style=""color:#F33;font-weight:600;"">[精华]</a> " End if%>
    <%If rs("istop")=1 Then Response.Write "<span style=""font-weight:600;"">[置顶]</span>" End if%>
    <%If rs("logid")>0 Then Response.Write "<span color=red>[<a href=""../go.asp?logid=" & rs("logid") & """ target=""_blank"">日志</a>]</font>" End if%>
    	<a href="../group.asp?gid=<%=rs("teamid")%>&pid=<%If rs("idepth")="1" Then  response.write rs("parentid")&"#a_"&rs("postid") Else response.write rs("postid")%>" target='_blank'><%=RemoveHtml(rs("topic"))%></a>
	</span>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="300"><a href="../go.asp?userid=<%=rs("userid")%>" target="_blank"><font color=#0d4d89><%=rs("author")%></font></a> 发表于<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;"><%
		Response.write rs("addtime") & "</span>　IP:<span style=""font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;color:#777;"">" &  rs("addip")
	%></span>
	</td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="40" align="center">
<%
        Response.write "<a href='m_post.asp?Action=Del&id=" & rs("postid") & "' onClick='return confirm(""确定要删除此日志吗？"");'>删除</a>&nbsp;"
%>
</td>
  </tr>
  <tr>
    <td align="center" valign="top"><span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("postid")%></span></td>
    <td colspan="3" valign="top" style="word-wrap: break-word; word-break: break-all;color:#555;"><%=Left(oblog.Filt_html(rs("content")),200)%></td>
  </tr>
  <tr>
    <td height="8"></td>
    <td colspan="4"></td>
  </tr>
<%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
rs.Close
Set rs=Nothing
%>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="140" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页所有内容</td>
            <td> <strong>操作：</strong>
              <input name="Action" type="radio" value="Del">
              删除&nbsp;&nbsp;
              &nbsp;&nbsp;
              <input type="submit" name="Submit" value="执行"> </td>
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