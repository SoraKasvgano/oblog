<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_blogstar")=False Then Response.Write "��Ȩ����":Response.End
dim rs, sql
dim id,UserSearch,Keyword,strField,douname
douname=oblog.filt_badstr(request("douname"))

keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
strField=Trim(Request("Field"))
UserSearch=Trim(Request("UserSearch"))
Action=Trim(Request("Action"))
id=Trim(Request("id"))

if UserSearch="" then
	UserSearch=0
else
	UserSearch=CLng(UserSearch)
end if
G_P_FileName="m_blogstar.asp?UserSearch=" & UserSearch
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
		<li class="main_top_left left">�� �� ֮ �� �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_blogstar.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>���ٲ��ң�</strong></td>
      <td width="687" height="30"><select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="0">���500������֮��</option>
          <option value="1">ͨ����˵Ĳ���֮��</option>
          <option value="2">δͨ����˵Ĳ���֮��</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="m_blogstar.asp">����֮�ǹ�����ҳ</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="m_blogstar.asp">
  <tr class="tdbg">
      <td width="120"><strong>�߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
	  <option value="blogname" selected>����֮����</option>
	  <option value="username" selected>�û���</option>
	  <option value="nickname" selected>�û��ǳ�</option>
      <option value="UserID" >����֮��ID</option>

      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10">
        ��Ϊ�գ����ѯ����</td>
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
select Case LCase(Action)
	Case "modify"
		call Modify()
	Case "savemodify"
		call SaveModify()
	Case "del"
		call DelUser()
	Case "pass0"
		Call Pass(0)
	Case "pass1"
		Call Pass(1)
	Case else
		call main()
end select
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	dim G_P_Guide
	G_P_Guide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='m_blogstar.asp'>����֮�ǹ���</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from oblog_blogstar order by id desc"
			G_P_Guide=G_P_Guide & "���500������֮��"
		case 1
			sql="select * from oblog_blogstar where ispass=1 order by id desc"
			G_P_Guide=G_P_Guide & "ͨ����˵Ĳ���֮��"
		case 2
			sql="select * from oblog_blogstar where ispass=0 order by id desc"
			G_P_Guide=G_P_Guide & "δͨ����˵Ĳ���֮��"
		case 10
			if Keyword="" then
				sql="select top 500 * from oblog_blogstar order by id desc"
				G_P_Guide=G_P_Guide & "���в���֮��"
			else
				select case strField
				case "UserID"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID������������</li>"
					else
						sql="select * from oblog_blogstar where id =" & CLng(Keyword)
						G_P_Guide=G_P_Guide & "����֮��ID����<font color=red> " & CLng(Keyword) & " </font>�Ĳ���֮��"
					end if
				case "blogname"
					sql="select * from oblog_blogstar where blogname like '%" & Keyword & "%' order by id  desc"
					G_P_Guide=G_P_Guide & "�������к��С� <font color=red>" & Keyword & "</font> ���Ĳ���֮��"
				case "username"
					sql="select * from oblog_blogstar where username like '%" & Keyword & "%' order by id  desc"
					G_P_Guide=G_P_Guide & "�û����к��С� <font color=red>" & Keyword & "</font> ���Ĳ���֮��"
				case "nickname"
					sql="select * from oblog_blogstar where usernickname like '%" & Keyword & "%' order by id  desc"
					G_P_Guide=G_P_Guide & "�������к��С� <font color=red>" & Keyword & "</font> ���Ĳ���֮��"
				end select
			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>����Ĳ�����</li>"
	end select
	G_P_Guide=G_P_Guide & "</td><td align='right'>"
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		G_P_Guide=G_P_Guide & "���ҵ� <font color=red>0</font> ������֮��</td></tr></table>"
		Response.write G_P_Guide
	else
    	G_P_AllRecords=rs.recordcount
		G_P_Guide=G_P_Guide & "���ҵ� <font color=red>" & G_P_AllRecords & "</font> ������֮��</td></tr></table>"
		Response.write G_P_Guide
		if G_P_This<1 then
       		G_P_This=1
    	end if
    	if (G_P_This-1)*G_P_PerMax>G_P_AllRecords then
	   		if (G_P_AllRecords mod G_P_PerMax)=0 then
	     		G_P_This= G_P_AllRecords \ G_P_PerMax
		  	else
		      	G_P_This= G_P_AllRecords \ G_P_PerMax + 1
	   		end if

    	end if
	    if G_P_This=1 then
        	showContent
        	Response.write oblog.showpage(true,true,"������֮��")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rs.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	Response.write oblog.showpage(true,true,"������֮��")
        	else
	        	G_P_This=1
           		showContent
           		Response.write oblog.showpage(true,true,"������֮��")
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i
    i=0
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�� �� ֮ �� �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="myform" method="Post" action="m_blogstar.asp" onsubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
<style type="text/css">
<!--
.border tr td {padding:3px 0!important;}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="title">
    <td align="center" width="44"><strong>ID</strong></td>
    <td align="center" width="100"><strong>����֮��ͼƬ</strong></td>
    <td align="center" width="120"><strong>���벩�� ����ʱ��</strong></td>
    <td align="center"><strong>����֮�Ǽ��</strong></td>
	  <td align="center" width="90"><strong>�û���/�ǳ�</strong></td>
    <td align="center" width="70"><strong>��˲���</strong></td>
    <td align="center" width="70"><strong>�������</strong></td>
  </tr>
          <%do while not rs.EOF %>
  <tr class="tdbg">
    <td align="center" style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("id")%></td>
    <td align="center">
	<a href="<%=rs("picurl")%>" target="_blank" title="����鿴��ͼ"><img src="<%=ProIco(rs("picurl"),1)%>" align="absmiddle" style="width:80px;height:60px;border:0;"></a>
	</td>
    <td>
	<span style="display:block;color:#666;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:12px;padding:0 0 0 8px!important;"><a href="<%=rs("userurl")%>" target="_blank" title="������ʸò���"><%=rs("blogname")%></a></span>
	<span style="display:block;color:#999;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:10px;padding:0 0 0 8px!important;">
	<%
	if rs("addtime")<>"" then
		Response.write rs("addtime")
	else
		Response.write "&nbsp;"
	end if
	%>
	</span>
	</td>
    <td valign="top"><span style="font-family:tahoma,Arial,Helvetica,sans-serif;padding:0 4px 0 4px!important;"><%=oblog.filt_html(rs("info"))%></span></td>
	<td align="center"><a href="<%=rs("userurl")%>" target="_blank" title="������ʸò���"><%=rs("username")&"<br/>"&rs("usernickname")%></a></td>
    <td align="center">
	<%
	select case rs("ispass")
		case 0
			Response.write "<span style=""color:#f30;font-weight:600;"">����</span>"
		case 1
			Response.write "<span style=""color:#090;font-weight:600;"">ͨ��</span>"
	end select
	%>&nbsp;
	<%
	If  rs("ispass")=0 Then
		Response.write "<a href='m_blogstar.asp?Action=pass1&id=" & rs("id") & "&douname="&rs("username")&"'>ͨ��</a>&nbsp;"
	Else
		Response.write "<a href='m_blogstar.asp?Action=pass0&id=" & rs("id") & "'>ȡ��</a>&nbsp;"
	End If
	%>
	</td>
    <td align="center">
<%
Response.write "<a href='m_blogstar.asp?Action=Modify&id=" & rs("id") & "'>�޸�</a>&nbsp;"
Response.write "<a href='m_blogstar.asp?Action=Del&id=" & rs("id") & "' onClick='return confirm(""ȷ��Ҫɾ���˲���֮����"");'>ɾ��</a>&nbsp;"
%>
	</td>
  </tr>
          <%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
%>
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


sub Modify()
	dim rsUser,sqlUser
	id=CLng(id)
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_blogstar where id=" & id
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���Ĳ���֮�ǣ�</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�޸Ĳ���֮����Ϣ</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<FORM name="Form1" action="m_blogstar.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td>blog����</td>
      <td><input name=blogname   type=text id="blogname" value="<%=rsuser("blogname")%>" size=30 maxlength=20></td>
    </tr>
    <TR class="tdbg" >
      <TD width="40%">���ӵ�ַ��</TD>
      <TD width="60%"> <INPUT name="userurl" value="<%=rsUser("userurl")%>" size=50   maxLength=250> <a href="<%=rsuser("userurl")%>" target="_blank">�鿴</a>
      </TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%"> ͼƬ����(<strong><font color="#FF0000">�뽫��ͼƬ�ֹ���Ϊ���ʵĳߴ�</font></strong>)��</TD>
      <TD width="60%"> <INPUT name=picurl value="<%=rsUser("picurl")%>" size=50 maxLength=250><a href="<%=rsuser("picurl")%>" target="_blank">�鿴</a></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">��飺</TD>
      <TD width="60%"><textarea name="bloginfo" cols="40" rows="5"><%=oblog.filt_html(rsuser("info"))%></textarea></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">״̬��</TD>
      <TD width="60%"><input type="radio" name="ispass" value=0 <%if rsUser("ispass")=0 then Response.write "checked"%>>
        δͨ�����&nbsp;&nbsp; <input type="radio" name="ispass" value=1 <%if rsUser("ispass")=1 then Response.write "checked"%>>
        ��ͨ�����</TD>
    </TR>
    <TR class="tdbg" >
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"> <input name="id" type="hidden" id="id" value="<%=rsUser("id")%>"></TD>
    </TR>
  </TABLE>
</form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%
	rsUser.close
	set rsUser=nothing
end sub


sub SaveModify()
	If Request.QueryString <>"" Then Exit Sub
	dim rsuser,sqlUser
	id=CLng(id)
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_blogstar where id=" & id
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	rsUser("blogname")=Trim(Request("blogname"))
	rsUser("userurl")=Trim(Request("userurl"))
	rsUser("picurl")=Trim(Request("picurl"))
	rsUser("ispass")=Trim(Request("ispass"))
	rsUser("info")=Trim(Request("bloginfo"))
	rsUser("addtime")=oblog.ServerDate(now())
	rsUser.update
	rsUser.Close
	set rsUser=Nothing
	WriteSysLog "�������޸Ĳ���֮�����ϲ�����Ŀ���û�ID��"&id&"",""
	oblog.ShowMsg "�޸ĳɹ�!",""
end sub

sub DelUser()
	id=CLng(id)
	oblog.execute("delete from oblog_blogstar where id="&id)
	WriteSysLog "������ɾ������֮�ǲ�����Ŀ���û�ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "ɾ���ɹ���",""
end sub

sub Pass(iState)
	id=CLng(id)
	oblog.execute("Update  oblog_blogstar Set ispass="& Cint(iState) &" where id="&id)
	If iState=0 Then
		WriteSysLog "������ȡ������֮�ǲ�����Ŀ���û�ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "��ȡ���ò���֮���ʸ�",""
	Else
		WriteSysLog "��������׼����֮�ǲ�����Ŀ���û�ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
		If int(oblog.CacheConfig(86)) = 1 Then 
		oblog.execute("INSERT INTO oblog_pm(incept,sender,topic,content) VALUES('"&doUname&"','ϵͳ����Ա','ϵͳ֪ͨ!����Ϊ��վ����֮��!','��ϲ,���Ѿ�����׼��Ϊ��վ���ٵĲ���֮��!�ٽ�����Ŷ!(����Ϣϵͳ�Զ�����,�Ķ��󽫱��Զ�ɾ��.�����ػظ�!)')")
		End If 
		oblog.ShowMsg "����׼�ò���֮�����룡",""
	End If
end Sub
Set oblog = Nothing
%>