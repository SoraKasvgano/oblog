<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
If CheckAccess("r_user_rblog")=False Then Response.Write "��Ȩ����":Response.End
dim rs, sql
dim id,cmd,Keyword,sField,sDate1,sDate2,blog
'-----------------------------
Dim Z_logRole,Z_classRole
	If oblog.CheckAdmin(1) Then
		Z_classRole=" "
		Else 
		Z_logRole=session("r_classes1")
		If Len(z_logrole) > 0 Or Not IsNull(z_logrole) Then
			If InStr(z_logrole,",") Then
				Z_classRole=" and classid in("&Z_logRole&") "
			ElseIf  Len(z_logrole) > 0 Then
				Z_classRole=" and classid = "&Int(Z_logRole)&" "
			End If
		End If
	End If
'-----------------------------
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
G_P_FileName="m_r_blog.asp?cmd=" & cmd & "&field=" & sField & "&keyword=" & keyword & "&date1=" & sDate1 & "&date2=" &sDate2
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
		<li class="main_top_left left">�� �� �� ־ �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form2" method="post" action="m_r_blog.asp">
  <tr class="tdbg">
      <td width="120"><strong>��ѯ��</strong></td>
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
        ��Ϊ�գ����ѯ����&nbsp;&nbsp;&nbsp;&nbsp;<a href="m_r_blog.asp?action=clear" onclick="if(confirm('ȷ��Ҫ�������վ�𣿲��ɻָ���')==false) return false;">ȫ�����</td>
  </tr>
</form>
  <form name="form3" method="post" action="m_r_blog.asp">
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
		If id="" Then
			oblog.ShowMsg "��ѡ��Ҫ��������־��","m_r_blog.asp"
			Response.End
		End If
		Set blog=New Class_blog
		Call blog.DeleteFiles(id,"")
		Set blog=Nothing
		oblog.execute("delete from oblog_log where logid In ("&id & ")")
		oblog.Execute("delete from oblog_comment where mainid in ("&id & ")")
		'��̬�ļ��Ѿ�ɾ��
		'ɾ�������ļ�
		WriteSysLog "��������־����ɾ��������Ŀ����־ID��"&id&"",oblog.NowUrl&"?"&OB_IIF(Request.QueryString,Request.Form)
		oblog.ShowMsg "ɾ���ɹ���",""
	Case "clear"
		Set blog=New Class_blog
		Call blog.DeleteFiles(id,"")
		Set blog=Nothing
		oblog.execute("delete from oblog_log where isdel=1")
		oblog.Execute("delete from oblog_comment where isdel=1")
		'ɾ�������ļ�
		WriteSysLog "��������ջ���վ����",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "����վ����ȫ��������ϣ�",""
	Case "renew"
		If id="" Then
			oblog.ShowMsg "��ѡ��Ҫ��������־��","m_r_blog.asp"
			Response.End
		End If
		oblog.execute("update oblog_log Set isdel=0 where logid In ("&id & ")")
'		oblog.execute("update oblog_comment Set isdel=0 where mainid In ("&id & ")")
		'��������ҳ��
		DoUpdatelog id
		WriteSysLog "��������־���·���������Ŀ����־ID��"&id&"",oblog.NowUrl&"?"&OB_IIF(Request.QueryString,Request.Form)
		Response.Redirect "m_r_blog.asp"
	Case Else
		call main()
end select
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	Dim sQryFields
	sQryFields="top 500 topic,logtext,logid,userid,addtime,passcheck,isbest,author,addip,classid"
	select case cmd
		case 10
			if Keyword="" then
				sql="select " & sQryFields & " from oblog_log  Where isdel=1 " & Z_classRole & " order by logid desc"
				sGuide=sGuide & "������־"
			else
				select case sField
				case "userid"
					if Not IsNumeric(Keyword) then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID������������</li>"
					else
						sql="select " & sQryFields & " from oblog_log where isdel=1  " & Z_classRole & "  and  userid =" & CLng(Keyword)
						sGuide=sGuide & "����ID����<font color=red> " & CLng(Keyword) & " </font>����־"
					end if
				case "author"
					sql="select " & sQryFields & " from oblog_log where isdel=1  and  author like '%" & Keyword & "%' " & Z_classRole & "  order by logid  desc"
					sGuide=sGuide & "���������к��С� <font color=red>" & Keyword & "</font> ������־"
				case "ip"
					sql="select " & sQryFields & " from oblog_log where isdel=1  and  addip like '%" & Keyword & "%' " & Z_classRole & "  order by logid  desc"
					sGuide=sGuide & "������־ʱ��IP�к��С� <font color=red>" & Keyword & "</font> ������־"
				case "title"
					sql="select " & sQryFields & " from oblog_log where isdel=1  and topic like '%" & Keyword & "%'  " & Z_classRole & " order by logid  desc"
					sGuide=sGuide & "��־�����к��С� <font color=red>" & Keyword & "</font> ������־"
				case "content"
					sql="select " & sQryFields & " from oblog_log where  isdel=1  and logtext like '%" & Keyword & "%'  " & Z_classRole & " order by logid  desc"
					sGuide=sGuide & "��־�����к��С� <font color=red>" & Keyword & "</font> ������־"
				end select
			end if
		Case 11
			sDate1=DeDateCode(sDate1)
			sDate2=DeDateCode(sDate2)
			If sDate1<>"" And sDate2<>"" Then
				sql="select " & sQryFields & " from oblog_log where truetime>=" & G_Sql_d_Char & sDate1 & G_Sql_d_Char & " And  truetime<=" & G_Sql_d_Char & sDate2 & G_Sql_d_Char &  " and isdel=1  " & Z_classRole & " order by logid  desc"
				sGuide=sGuide & "ʵ�ʷ���ʱ���� <font color=red>" & sDate1 & "</font> �� <font color=red>" & sDate2 & "</font> ����־"
			End If
		case else
			sql="select " & sQryFields & " from oblog_log  Where isdel=1  " & Z_classRole & " order by logid desc"
			sGuide=sGuide & "������־"
	end select
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	If Trim(Sql)="" Then
		oblog.ShowMsg "��������ȷ�Ĳ�ѯ������",""
	End If
	'Response.Write sql
	rs.Open sql,Conn,1,1
  Call oblog.MakePagebar(rs,"ƪ��־")
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
  <form name="myform" method="Post" action="m_r_blog.asp" onsubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
          <%do while not rs.EOF %>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="0" style="word-wrap: break-word; word-break: break-all;">
  <tr>
    <td align="center" style="background:#B3D1EA;border-bottom:1px #000 dotted;" width="30">
    	<input name='id' type='checkbox' onclick="unselectall()" id="id" value='<%=cstr(rs("logid"))%>'>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;"><span>[<%=oblog.GetClassName(2,0,rs("classid"))%>]</span>
    	<a href="../go.asp?logid=<%=rs("logid")%>" target="_blank" style="margin:0 0 0 10px;color:#333;"><%=RemoveHtml(rs("topic"))%></a>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="290"><font color=#0d4d89><%=rs("author")%></font> ������
	<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;">
	<%
		Response.write rs("addtime") & "</span>��<span style=""font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;color:#777;"">IP:" &  rs("addip")
	%></span>
	</td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="60">
<%			Response.write "<a href='m_r_blog.asp?Action=renew&id=" & rs("logid") & "' onClick='return confirm(""ȷ��Ҫ�ָ�����־��"");'>�ָ�</a>&nbsp;"
        Response.write "<a href='m_r_blog.asp?Action=del&id=" & rs("logid") & "' onClick='return confirm(""ȷ��Ҫɾ������־��"");'>ɾ��</a>&nbsp;"
%>
</td>
  </tr>
  <tr>
    <td align="center" valign="top"><span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("logid")%></span></td>
    <td colspan="3" valign="top" style="word-wrap: break-word; word-break: break-all;color:#555;"><%=Left(oblog.Filt_html(RemoveHtml(rs("logtext"))),200)%></td>
  </tr>
  <tr>
    <td height="8" colspan="4"></td>
  </tr>
</table>
<%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
rs.Close
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="140" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ѡ�б�ҳ������־</td>
            <td> <strong>������</strong>
              <input name="Action" type="radio" value="del">
              ����ɾ��&nbsp;&nbsp;
              <input name="Action" type="radio" value="renew">
              �ָ�&nbsp;&nbsp;
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
end sub
'������־
Sub DoUpdatelog(ids)
    Server.ScriptTimeOut = 999999999
    Dim  rs
	Dim sScore
    Set rs = oblog.execute("select userid,logid,subjectid,classid,scores from oblog_log where logid in (" & ids & ")")
    Set blog = New class_blog
    Do While Not rs.Eof
		Call OBLOG.log_count(rs(0),rs(1),rs(2),rs(3),"+")
		sScore=rs(4)+CLng(oblog.CacheScores(4))
		If IsNull(sScore) Then sScore = CLng(oblog.CacheScores(4))
		oblog.GiveScore "",sScore,rs(0)
		oblog.execute("update oblog_comment Set isdel=0 where mainid In ("&rs(1) & ")")
        blog.userid = rs(0)
        blog.update_log rs(1), 0
        blog.update_index  0
        rs.movenext
    Loop
    Set rs = Nothing
End Sub
Set oblog = Nothing
%>