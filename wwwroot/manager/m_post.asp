<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_group_blog")=False Then Response.Write "��Ȩ����":Response.End
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
<title>oBlog--��̨����</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">Ⱥ �� �� �� �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_post.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>���ٲ��ң�</strong></td>
      <td width="687" height="30"><a href="m_post.asp">��������</a>|&nbsp;&nbsp;<a href="m_post.asp?cmd=1">�����б�</a>|&nbsp;&nbsp;<a href="m_post.asp?cmd=3">�����б�</a>|&nbsp;&nbsp;<a href="m_post.asp?cmd=2">�ظ��б�</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="m_post.asp">
  <tr class="tdbg">
      <td width="120"><strong>�߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
	      <option value="author" selected>�û�����</option>
	      <option value="userid" >�û�ID</option>
	      <option value="ip" selected>����IP</option>
	      <option value="title" >��������</option>
	      <option value="content" >��������</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="cmd" type="hidden" id="cmd" value="10">
        ��Ϊ�գ����ѯ����</td>
  </tr>
</form>
  <form name="form3" method="post" action="m_post.asp">
  <tr class="tdbg">
      <td width="120"><strong>��ʱ�����β�ѯ��</strong></td>
    <td>
    	��ʼʱ�䣺<input type="text" name="date1" size=12 maxlength=10>
    	����ʱ�䣺<input type="text" name="date2" size=12 maxlength=10>

      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="cmd" type="hidden" id="cmd" value="11">
      <br/>
        ʱ���ʽ��YYYYMMDDHHmm����2006��6��6��9�㣬������2006060609,������ʽ����֧��</td>
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
		WriteSysLog "������ɾ��"&oblog.CacheConfig(69)&"���������Ŀ��"&oblog.CacheConfig(69)&"ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "ɾ���ɹ���",""
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
			sGuide=sGuide & "���500ƪ����"
		Case 1
			sql="select " & sQryFields &" from oblog_teampost Where idepth=0 order by postid desc"
		Case 2
			sql="select " & sQryFields &" from oblog_teampost Where idepth>0 order by postid desc"
		Case 3
			sql="select " & sQryFields &" from oblog_teampost Where idepth=0 And isbest=1 order by postid desc"
		case 10
			if Keyword="" then
				sql="select " & sQryFields &" from oblog_teampost order by postid desc"
				sGuide=sGuide & "������־"
			else
				select case sField
				case "userid"
					if Not IsNumeric(Keyword) then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID������������</li>"
					else
						sql="select " & sQryFields & " from oblog_teampost where userid =" & CLng(Keyword)
						sGuide=sGuide & "����ID����<font color=red> " & CLng(Keyword) & " </font>����־"
					end if
				case "author"
					sql="select " & sQryFields & " from oblog_teampost where author like '%" & Keyword & "%' order by postid  desc"
					sGuide=sGuide & "���������к��С� <font color=red>" & Keyword & "</font> ������־"
				case "ip"
					sql="select " & sQryFields & " from oblog_teampost where addip like '%" & Keyword & "%' order by postid  desc"
					sGuide=sGuide & "������־ʱ��IP�к��С� <font color=red>" & Keyword & "</font> ������־"
				case "title"
					sql="select " & sQryFields & " from oblog_teampost where topic like '%" & Keyword & "%' order by postid  desc"
					sGuide=sGuide & "�����к��С� <font color=red>" & Keyword & "</font> ������־"
				case "content"
					sql="select " & sQryFields & " from oblog_teampost where content like '%" & Keyword & "%' order by postid  desc"
					sGuide=sGuide & "�����к��С� <font color=red>" & Keyword & "</font> ������־"
				end select
			end if
		Case 11
			sDate1=DeDateCode(sDate1)
			sDate2=DeDateCode(sDate2)
			If sDate1<>"" And sDate2<>"" Then
				sql="select " & sQryFields & " from oblog_teampost where addtime>=" & G_Sql_d_Char & sDate1 & G_Sql_d_Char & " And  addtime<=" & G_Sql_d_Char & sDate2 & G_Sql_d_Char &  " order by postid  desc"
				sGuide=sGuide & "ʵ�ʷ���ʱ���� <font color=red>" & sDate1 & "</font> �� <font color=red>" & sDate2 & "</font> ������"
			End If
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>����Ĳ�����</li>"
	end select
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'Response.Write sql
	rs.Open sql,Conn,1,1
  Call oblog.MakePagebar(rs,"ƪ")
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
  <form name="myform" method="Post" action="m_post.asp" onsubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
          <%do while not rs.EOF %>
  <tr>
    <td align="center" style="background:#B3D1EA;border-bottom:1px #000 dotted;" width="30">
    	<input name='id' type='checkbox' onclick="unselectall()" id="id" value='<%=cstr(rs("postid"))%>'>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;">
    <span style="margin:0 0 0 10px;">
	<%If rs("idepth")=0 Then Response.Write "<a href=""m_post.asp?cmd=1"" style=""color:#217DBD;font-weight:600;"">[����]</a> " Else Response.Write "<a href=""m_post.asp?cmd=2"" style=""color:#99F;font-weight:600;"">[�ظ�]</a> " End if%>
    <%If rs("isbest")=1 Then Response.Write "<a href=""m_post.asp?cmd=3"" style=""color:#F33;font-weight:600;"">[����]</a> " End if%>
    <%If rs("istop")=1 Then Response.Write "<span style=""font-weight:600;"">[�ö�]</span>" End if%>
    <%If rs("logid")>0 Then Response.Write "<span color=red>[<a href=""../go.asp?logid=" & rs("logid") & """ target=""_blank"">��־</a>]</font>" End if%>
    	<a href="../group.asp?gid=<%=rs("teamid")%>&pid=<%If rs("idepth")="1" Then  response.write rs("parentid")&"#a_"&rs("postid") Else response.write rs("postid")%>" target='_blank'><%=RemoveHtml(rs("topic"))%></a>
	</span>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="300"><a href="../go.asp?userid=<%=rs("userid")%>" target="_blank"><font color=#0d4d89><%=rs("author")%></font></a> ������<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;"><%
		Response.write rs("addtime") & "</span>��IP:<span style=""font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;color:#777;"">" &  rs("addip")
	%></span>
	</td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="40" align="center">
<%
        Response.write "<a href='m_post.asp?Action=Del&id=" & rs("postid") & "' onClick='return confirm(""ȷ��Ҫɾ������־��"");'>ɾ��</a>&nbsp;"
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
              ѡ�б�ҳ��������</td>
            <td> <strong>������</strong>
              <input name="Action" type="radio" value="Del">
              ɾ��&nbsp;&nbsp;
              &nbsp;&nbsp;
              <input type="submit" name="Submit" value="ִ��"> </td>
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