<!--#include file="inc/inc_sys.asp"-->
<%If CheckAccess("r_user_news")=False Then Response.Write "��Ȩ����":Response.End%>
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
		<li class="main_top_left left">oBlogվ�ڶ��Ź�����ҳ</li>
		<li class="main_top_right right"> </li>
	</ul>
		<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" id="table1">
  <form name="form1" action="m_pmall.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>���ٲ��ң�</strong></td>
      <td width="687" height="30">
		<select size=1 name="cmd" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="0">����Ա������վ�ڶ���</option>

        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="m_pmall.asp?cmd=0">վ�ڶ��Ź�����ҳ</a></td>
    </tr>
  </form>
<!--   <form name="form2" method="post" action="m_pmall.asp">
  <tr class="tdbg">
    <td width="120"><strong>�߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
      <option value="id" selected="selected">����ID</option>
	  <option value="topic">���ű���</option>
      <option value="group" >�û���ID</option>
	  <option value="sender" >������</option>
	  <option value="incept" >�ռ���</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="cmd" type="hidden" id="cmd" value="10">
	  ��Ϊ�գ����ѯ����վ�ڶ���</td>
  </tr>
</form> -->
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%
Dim rst
Dim cmd,sql,rs,Keyword,sField
Set rst=oblog.Execute("select groupid,g_name From oblog_groups Order By groupid")
action=Request("action")
cmd=Trim(Request("cmd"))
sField=Trim(Request("Field"))
keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if

G_P_FileName="m_pmall.asp?cmd=" & cmd
if sField<>"" then
	G_P_FileName=G_P_FileName&"&Field="&sField
end if
if keyword<>"" then
	G_P_FileName=G_P_FileName&"&keyword="&keyword
	cmd=10
End If
If keyword<>"" Or cmd<>"" Then call main():Response.End
If action = "del" Then Call delone():Response.End
call send()
select case action
	case "save"
	call save()
end select
sub send()
	dim rs
%>
<SCRIPT language=javascript>
function changincept()
{
	document.oblogform.incept.value = document.oblogform.selectincept.value;
}
</SCRIPT>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">վ���������Ͷ���</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table align="center" cellpadding="1" cellspacing="1" Class="border">
<form action="m_pmall.asp?action=save" method="post" name="oblogform">
  <tr class="tdbg">
    <td>Ŀ���û���</td>
    <td>
      		<%
      		rst.Movefirst
      		Do While Not rst.Eof
      			%>
      		<input type="checkbox" name="groupid" value="<%=rst(0)%>"><%=rst(1)%><br/>
      			<%
      			rst.MoveNext
      		Loop
      		%>
	</td>
  </tr>
 <tr class="tdbg">
    <td>���⣺</td>
	<td><input type="text" name="topic" size="45" maxlength="50" /></td>
 </tr>
   <tr class="tdbg">
    <td>���ݣ�<br />(250����)</td>
	<td><textarea name="content" cols="45" rows="8"></textarea></td>
  </tr>
 <tr class="tdbg">
   <td></td><td> <INPUT type="hidden" name="id" value="">
        <input type="submit"  value=" �ύ ">
		��<br></td>
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
end sub

sub save()
	dim incept,content,sql,rs,inceptid,topic,username,sqlt,rst,u,s
	content=Trim(Request("content"))
	topic=Trim(Request("topic"))
	u=Replace(Request("groupid")," ","")
	if content="" then Response.write("<font color=red>���󣺶���Ϣ���ݲ���Ϊ��</font><br />")
	if topic="" then Response.write("<font color=red>���󣺶���Ϣ���ⲻ��Ϊ��</font><br />")
	if u="" then Response.write("<font color=red>��������ѡ��һ��Ŀ����</font><br />")
	If content<>"" And topic<>"" And u<>"" Then
		sqlt="select * from oblog_pm"
		set rs=Server.CreateObject("adodb.recordset")
		rs.open sqlt,conn,1,3
		rs.addnew
			rs("incept")="0"
			rs("topic")=oblog.Interceptstr(topic,100)
			rs("content")=oblog.Interceptstr(content,250)
			rs("issys")=1
			rs("groups")=u
			rs("sender")="����Ա"
		rs.update
		rs.close
		set rs=Nothing
		WriteSysLog "������վ�ڶ��Ź��������Ŀ���û���ΪID��"&u&"",oblog.NowUrl&"?"&Request.QueryString
		Response.Write("<ul><li>����Ϣ���ͳɹ�</li></ul>")
	end if
end Sub

sub main()
	if cmd="" then
		cmd=0
	else
		cmd=CLng(cmd)
	end If
	sGuide=""
	select case cmd
		case 0
			sql="select top 500 * from oblog_pm Where issys=1 order by id desc"
			sGuide=sGuide & "����Ա������վ�ڶ���"
		case 10
'			if Keyword="" then
'				sql="select top 500 * from oblog_pm order by id Desc"
'				sGuide=sGuide & "����վ�ڶ���"
'			else
'				select case LCase(sField)
'				case "id"
'					if IsNumeric(Keyword)=false then
'						FoundErr=true
'						ErrMsg=ErrMsg & "<br><li>ID������������</li>"
'					else
'						sql="select * from oblog_pm where id =" & CLng(Keyword)
'						sGuide=sGuide &  "ID����<font color=red> " & CLng(Keyword) & " </font>��վ�ڶ���"
'					end if
'				case "group"
'					sql="select * from oblog_pm where  groups like '"&Keyword&",%' or groups like '%,"&Keyword&"' or groups like '%,"&Keyword&",%' or groups ='"&oblog.l_uGroupId&"'"
'					sGuide=sGuide & "Ŀ���û���IDΪ�� <font color=red>" & Keyword & "</font> ����վ�ڶ���"
'				case "sender"
'					sql="select * from oblog_pm where sender = '" & Keyword & "'"
'					sGuide=sGuide &"������Ϊ�� <font color=red>" & Keyword & "</font> ����վ�ڶ���"
'				case "incept"
'					sql="select * from oblog_pm where incept='" & Keyword&"'"
'					sGuide=sGuide &"�ռ���Ϊ�� <font color=red>" & Keyword & "</font> ����վ�ڶ���"
'				end select
'			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>����Ĳ�����</li>"
	end select
	If sGuide="" Then sGuide="վ�ڶ��Ź���"
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
'	Response.write sql
	rs.Open sql,Conn,1,1
	%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left"><%=sGuide%></li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
			<%
	Call oblog.MakePageBar(rs,"��վ�ڶ���")
	%>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
	<%
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i
    i=0
%>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border" id="table3">
          <tr class="title">
            <td width="42" align="center"><strong>ID</strong></td>
            <td width="77" height="22" align="center"><strong>������</strong></td>
            <td width="97" height="22" align="center"><strong>�ռ���</strong></td>
            <td height="22" width="93" align="center"><strong>����</strong></td>
            <td height="30" align="center"><strong>����</strong></td>
            <td width="108" align="center"><strong>ʱ��</strong></td>
            <td  width="36" height="22"  align="center" ><strong>����</strong></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
            <td width="42" align="center"><%=rs("id")%></td>
            <td width="77" align="center">
			<%
			If rs("issys")= 1 Then
				Response.Write "<font color=red style=font-weight:600>" &rs("sender") &"</font>"
			Else
				Response.Write oblog.filt_html(rs("sender"))
			End If%>
			</td>
            <td width="97" align="center"><%
				If rs("incept")="0" Then
					Response.Write "<font color=green style=font-weight:600>" &GetGroupName(rs("groups"))& "</font>"
				Else
					Response.Write oblog.filt_html(rs("incept"))
				End if%>
				</td>
            <td align="center"  width="93">
            <%=rs("topic")
            %>
            </td>
        	<td width="282" align="center">
            	 <%=Left(rs("content"),50)%>
		    </td>
            <td align="center"> <%=rs("addtime")%>
		    </td>
            <td  align="center" width="36">
        <a href="m_pmall.asp?action=del&id=<%=rs("id")%>" onClick="return confirm('ȷ��Ҫɾ����');"> ɾ��</a>            </td>
          </tr>
          <%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext

Loop
Response.write "</table>"
end Sub

sub delone()
	Dim id
	id=CLng(Request("id"))
	oblog.execute("DELETE FROM oblog_pm WHERE id= " & id)
	WriteSysLog "������վ�ڶ���ɾ��������Ŀ�����ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "�����ɹ�","m_pmall.asp?cmd=0"
end Sub

Function GetGroupName(groupid)
	Dim rs,tstr,i
	If InStr(groupid,",")<0 Then
		Set rs = oblog.execute ("select g_name FROM oblog_groups WHERE groupid = " & groupid)
		If Not rs.EOF Then
			GetGroupName = rs(0)
		Else
			GetGroupName = "�û��鲻���ڻ����Ѿ���ɾ����"
		End  if
	Else
		groupid = Split (groupid,",")
		For i= 0 To UBound (groupid)
			Set rs = oblog.execute ("select g_name FROM oblog_groups WHERE groupid = " & groupid(i))
			If Not( rs.eof Or rs.bof) Then tstr = tstr & "," & rs(0)
		Next
		GetGroupName = Replace (tstr,",","",1,1,1)
	End if
End Function
Set oblog = Nothing
%>
</body>
</html>