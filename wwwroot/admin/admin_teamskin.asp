<!--#include file="inc/inc_sys.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--<%=oblog.CacheConfig(69)%> ģ �� �� ��</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.form2.chkAll.checked){
	document.form2.chkAll.checked = document.form2.chkAll.checked&0;
    }
}

function checkAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
</script>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">

<%
dim action,ispass,rstClass,sClasses
Action=Trim(Request("Action"))
Set rstClass=Server.CreateObject("Adodb.RecordSet")
rstClass.Open "select * From oblog_skinclass Where iType=1",conn,1,3
If Not rstClass.Eof Then
	Do While Not rstClass.Eof
		sClasses= sClasses & "<option value=" & rstClass("classid") & " >" & rstClass("classname") & "(" & rstClass("icount") & ")</option>" & vbcrlf
		rstClass.MoveNext
	Loop
	rstClass.MoveFirst
End if
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left"><%=oblog.CacheConfig(69)%> ģ �� �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="tdbg">
      <td width="100" height="30"><strong>�������ӣ�</strong></td>
      <td  width="687" height="30">
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_teamskin.asp?action=skinclass">ģ�����ά��</a>&nbsp;|&nbsp;<a href="admin_teamskin.asp?action=showskin&ispass=1">��ͨ����˵�ģ��</a>&nbsp;|&nbsp;<a href="admin_teamskin.asp?action=showskin&ispass=0">δͨ����˵�ģ��</a></td>
    </tr>
  <form name="form1" action="admin_teamskin.asp?action=showskin&ispass=1" method="post">
    <tr class="tdbg">
      <td width="100" height="30"><strong>��������ˣ�</strong></td>
      <td width="687" height="30">
      	<select size=1 name="classid">
      	  <option value="0">------��δ����------</option>
          <%=sClasses%>
        </select>
        <input type="submit" value=" �� �� "></td>
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
select case Action
	case "saveconfig"
		call saveconfig()
	case "showskin"
		call showskin()
	case "modiskin"
		call modiskin()
	case "savedefault"
		call savedefault()
	case "delconfig"
		call delconfig()
	case "addskin"
		call addskin()
	case "saveaddskin"
		call saveaddskin()
	case "passskin"
		call passskin
	case "unpassskin"
		call unpassskin
	case "move"
		call SaveMove()
	Case "skinclass"
		Call ShowClasses
	Case "saveclass"
		Call SaveClass
	Case "delclass"
		Call DelClass
end select

sub showskin()
	dim rs,psql,sql,classid
ispass=CLng(Request("ispass"))
classid=Request("classid")
If classid<>"" Then Classid=Int(classid)
if ispass=1 Then
	G_P_FileName="admin_teamskin.asp?action=showskin&ispass=1&classid="&Classid
	psql=" where ispass=1 "
else
	G_P_FileName="admin_teamskin.asp?action=showskin&ispass=0&classid="&Classid
	psql=" where ispass=0 "
end if

If classid<>"" Then
	If classid=0 Then
		psql=" where ispass=1 And (classid=0 Or classid Is Null) "
	Else
		psql=" where ispass=1 And classid=" & classid
	End If
End If

	if Request("page")<>"" then
	    G_P_This=cint(Request("page"))
	else
		G_P_This=1
	end if
	set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select id,userskinname,skinauthor,skinauthorurl,isdefault,ispass,skinpic,classid from oblog_teamskin "&psql&" order by id desc "
'Response.Write Sql
'Response.End
		rs.Open sql,Conn,1,1
	  	if rs.eof  then
'			showContent(rs)
			G_P_Guide=G_P_Guide & " (����0��ģ��)</h1>"
			Response.write "<div align='center'>"&G_P_Guide&"</div>"
		else
	    	G_P_AllRecords=rs.recordcount
			G_P_Guide=G_P_Guide & " (����" & G_P_AllRecords & "��ģ��)</h1>"
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
	        	Call showContent(rs)
	        	Response.write oblog.showpage(true,true,"��ģ��")
	   	 	else
	   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
	         	   	rs.move  (G_P_This-1)*G_P_PerMax
	         		dim bookmark
	           		bookmark=rs.bookmark
	        	else
		        	G_P_This=1
		    	end if
		    	Call showContent(rs)
		    	Response.write oblog.showpage(true,true,"��ģ��")
			end if
		end if
end sub

sub showContent(rs)
	dim i
	i=0
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left"><%if ispass=1 then Response.Write "ͨ����˵�ģ��" else Response.write "δͨ����˵�ģ��"%></li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form name="form2" method="post" action="admin_teamskin.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
  <tr class="topbg">
    <td align="center"><strong>ѡ��</strong></td>
    <td align="center"><strong>ģ�����</strong></td>
    <td align="center"><strong>ģ������</strong></td>
    <td><strong>����</strong></td>
  </tr>
    <%
do while not rs.eof
dim userskinname
    userskinname=rs("userskinname")
%>
  <tr class="topbg">
    <td width="30" valign="top">
	<div align="center"><input name="checkbox" type="checkbox" onClick="unselectall()" id= "checkbox" class="tdbg" value='<%=rs("id")%>'></div>
	<div align="center"><%= rs("id") %></div></td>
    <td width="120"><div align="center">
      	<%
      	Dim ClassId1
      	ClassId1=Ob_IIF(rs("classid"),0)
      	If ClassId1>0 Then
      		rstClass.Filter="classid=" & ClassId1
      		If Not rstClass.Eof Then
      			Response.Write rstClass("classname")
      		Else
      			Response.Write "--"
      		End If
      	Else
      		Response.Write "--"
      	End If
      	 %></div></td>
    <td width="140"><div align="center">
	  <a href="../showskin.asp?teamskinid=<%=rs("id")%>" target="_blank"><img style="width:120px;height:77px;border:1px #888 solid;" src="<%=ProIco(rs("skinpic"),3)%>" /><br />
	  <%if rs("isdefault")=1 then
	  Response.Write "<font style=""color:#f00;font-weight:600;"">Ĭ��ģ�壺"&userskinname&"</red>"
	  else
	  Response.Write userskinname
	  end if
	  %></a>
	  </div></td>
    <td><div>
	  <strong>ģ�����ߣ�</strong><%if rs("skinauthorurl")="" or isnull(rs("skinauthorurl")) then
	  Response.Write rs("skinauthor")
	  else
	  Response.Write "<a href="""&oblog.filt_html(rs("skinauthorurl"))&""" target='_blank'>"&rs("skinauthor")&"</a>"
	  end if%>
	  </div>
	  <div><%if rs("ispass")=1 then Response.Write("<span style=""color:#317531;font-weight:600;"">�����</span>") else Response.Write("<span style=""color:#F30;font-weight:600;"">δ���</span>")%>������	<%if ispass=0 then%>
	<a href="admin_teamskin.asp?action=passskin&id=<%=rs("id")%>">ͨ�����</a>
	<%else%>
	<a href="admin_teamskin.asp?action=unpassskin&id=<%=rs("id")%>">ȡ�����</a>
	<%end if%></div>
	  <div><a href="../admin_edit.asp?action=modiskin&skintype=team&t=0&editm=1&skinorder=0&id=<%=rs("id")%>"  target="_blank">�޸���ģ��</a>��<a href="../admin_edit.asp?action=modiskin&skintype=team&t=0&editm=1&skinorder=1&id=<%=rs("id")%>"  target="_blank">�޸ĸ�ģ��</a>��<a href="admin_teamskin.asp?action=modiskin&id=<%=rs("id")%>">�޸�ģ��(�ı���ʽ)</a></div>
	  <div><a href="admin_teamskin.asp?action=delconfig&id=<%=rs("id")%>" style="color:#f00;font-weight:600;" onclick=return(confirm("ȷ��Ҫɾ�����ģ����"))>ɾ��ģ��</a></div>
	  </td>
  </tr>
    <%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
%>
    <tr>
      <td height="40" colspan="4" align="center" class="tdbg"><div align="center">
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=checkAll(this.form) value="checkbox" />
	  ȫѡ
	 <input type="radio" value="savedefault" name="action" checked>Ĭ��ģ��
	 <%if ispass=0 then%>
	  <input type="radio" value="passskin" name="action" >ͨ�����
	  <%else%>
	  <input type="radio" value="unpassskin" name="action">ȡ�����
	  <%end if%>
	   <input type="radio" value="delconfig" name="action" >ɾ��
	   <input type="radio" value="move" name="action" >�ƶ�����
	   <select name="classid">
	   	<%=sClasses%>
	  </select>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="submit" name="Submit" value="��������">
        </div></td>
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
	set rs=nothing
end sub

sub savedefault()
	If Request.QueryString <>"" Then Exit Sub
	dim isdefaultID
	isdefaultID=Trim(Request("checkbox"))
		if instr(isdefaultID,",")>0 then
		Response.Write("<script language=javascript>alert('�û�Ĭ��ģ��ֻ����ѡ��һ����');history.back();</script>")
		Response.End()
	elseif isdefaultID="" then
		Response.Write("<script language=javascript>alert('��ָ��Ҫ�趨ΪĬ�ϵ�ģ�壡');history.back();</script>")
		Response.End()
		exit sub
	end if
	oblog.execute("update oblog_teamskin set isdefault=0")
	oblog.execute("update oblog_teamskin set isdefault=1 where id="&isdefaultID)
	EventLog "�������趨Ĭ��Ⱥ��ģ�������Ŀ��ģ��ID��"&isdefaultID&"",""
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""�޸ĳɹ���"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub passskin()
	dim id
	id=Trim(Request("checkbox"))
	if instr(id,",")>0 then
	id=Replace(id," ","")
	oblog.execute("update oblog_teamskin set ispass=1 where id in ("&id&")")
	elseif id="" then
	id=CLng(Request("id"))
	oblog.execute("update oblog_teamskin set ispass=1 where id="&id)
	else
    oblog.execute("update oblog_teamskin set ispass=1 where id="&id)
	end if
	ReCountSkins
	EventLog "������ͨ�����Ⱥ��ģ�������Ŀ��ģ��ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "ͨ����˳ɹ�",""
end sub

sub unpassskin()
	dim id
	id=Trim(Request("checkbox"))
	if instr(id,",")>0 then
	id=Replace(id," ","")
	oblog.execute("update oblog_teamskin set ispass=0 where id in ("&id&")")
	elseif id="" then
	id=CLng(Request("id"))
	oblog.execute("update oblog_teamskin set ispass=0 where id="&id)
	else
	oblog.execute("update oblog_teamskin set ispass=0 where id="&id)
	end if
	ReCountSkins
	EventLog "������ȡ�����Ⱥ��ģ�������Ŀ��ģ��ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "ȡ����˳ɹ�",""
end sub


sub saveconfig()
	dim rs,sql
	if Trim(Request("userskinname"))="" then oblog.sys_err("ģ��������Ϊ��"):Response.End()
	if Trim(Request("skinmain"))="" then oblog.sys_err("��ģ�岻��Ϊ��"):Response.End()
	if Trim(Request("skinshowlog"))="" then oblog.sys_err("��ģ�岻��Ϊ��"):Response.End()
	set rs=Server.CreateObject("adodb.recordset")
	sql="select * from oblog_teamskin where id="&CLng(Request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs("userskinname")=Trim(Request("userskinname"))
	rs("skinauthor")=Trim(Request("skinauthor"))
	rs("skinmain")=Request("skinmain")
	rs("skinshowlog")=Request("skinshowlog")
	rs("skinpic")=Trim(Request("skinpic"))
	rs("classid")=Trim(Request("classid"))
	rs("skinauthorurl")=Trim(Request("skinauthorurl"))
	rs("isdefault") = 0
	rs("ispass") = 1
	rs.update
	rs.close
	set rs=nothing
	ReCountSkins
	EventLog "�������޸�Ⱥ��ģ��������ı���ʽ����Ŀ��ģ��ID��"&Request.QueryString("id")&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "����ɹ�",""
end sub

sub delconfig()
    dim id
	id=Trim(Request("checkbox"))
	if instr(id,",")>0 then
	id=Replace(id," ","")
	oblog.execute("delete from oblog_teamskin where id in ("&id&")")
	elseif id="" then
	id=CLng(Request.QueryString("id"))
		oblog.execute("delete from oblog_teamskin where id="&id)
	else
		oblog.execute("delete from oblog_teamskin where id="&id)
	end if
	ReCountSkins
	EventLog "������ɾ��Ⱥ��ģ�������Ŀ��ģ��ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "ɾ���ɹ�",""
end sub
sub modiconfig()
	dim rs
	set rs=oblog.execute("select * from oblog_teamskin where id="&CLng(Request.QueryString("id")))
End Sub
sub saveaddskin()
	dim rs,sql
	set rs=Server.CreateObject("adodb.recordset")
	if Trim(Request("userskinname"))="" then oblog.sys_err("ģ��������Ϊ��"):Response.End()
	if Trim(Request("skinmain"))="" then oblog.sys_err("��ģ�岻��Ϊ��"):Response.End()
	if Trim(Request("skinshowlog"))="" then oblog.sys_err("��ģ�岻��Ϊ��"):Response.End()
	sql="select * from oblog_teamskin where id="&CLng(Request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs.addnew
	rs("userskinname")=Trim(Request("userskinname"))
	rs("skinauthor")=Trim(Request("skinauthor"))
	rs("skinmain")=Trim(Request("skinmain"))
	rs("skinshowlog")=Trim(Request("skinshowlog"))
	rs("skinpic")=Trim(Request("skinpic"))
	rs("classid")=Trim(Request("classid"))
	rs("skinauthorurl")=Trim(Request("skinauthorurl"))
	rs("isdefault") = 0
	rs("ispass") = 1
	rs.update
	rs.close
	set rs=nothing
	ReCountSkins
	EventLog "���������Ⱥ��ģ�����",oblog.NowUrl&"?"&Request.QueryString
	Response.Redirect "admin_teamskin.asp?action=showskin&ispass=1"
end sub

sub modiskin()
	dim rs
	set rs=oblog.execute("select * from oblog_teamskin where id="&CLng(Request.QueryString("id")))
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�޸��û�ģ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
    <tr class="tdbg">
    <td width="253" height="30"><strong>�����޸ĵ�ģ���ǣ�<%=rs("userskinname")%></strong></td>
    <td width="516" height="30">
	<a href="admin_teamskin.asp?action=modiskin&id=<%=rs("id")%>">�޸�ģ��</a>����<a href="admin_teamskin.asp?action=showskin&ispass=1">���ع���˵�</a>
      <a href="admin_skin_help.asp" target="_blank"><strong>ģ���ǰ���</strong></a></td>
    </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�޸�ģ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="admin_teamskin.asp?id=<%=CLng(Request.QueryString("id"))%>" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td height="25" class="tdbg">ģ�����ƣ�
        <input name="userskinname" type="text" id="userskinname" value=<%=rs("userskinname")%>>
        �������ߣ�
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>>
         �趨����:
        <select size=1 name="classid">
      	  <option value="0" <%If OB_IIF(rs("classid"),0)=0 Then Response.Write " Selected" End If%>>------��δ����------</option>
      	  <%
	    Do While Not rstClass.Eof
	    	%>
			<option value="<%=rstClass("classid")%>" <%If CLng(OB_IIF(rstClass("classid"),0))=rs("classid") Then Response.Write " Selected" End If%>><%=rstClass("classname")%>(<%=rstClass("icount")%>)</option>
			<%
			rstClass.MoveNext
		Loop%>
        </select>
        <br>
        �������ӣ�
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="40" value="<%=rs("skinauthorurl")%>">
         <br>
        Ԥ��ͼƬ<strong>��
        <input name="skinpic" type="text" id="skinpic" size="40" value="<%=rs("skinpic")%>">
        </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <strong>��ģ�壺</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"><%if rs("skinmain")<>"" then Response.Write Server.HtmlEncode(rs("skinmain")) else Response.Write("")%></textarea>
        <br>
        <br>
        <strong>��ģ�壺 <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"><%if rs("skinshowlog")<>"" then Response.Write Server.HtmlEncode(rs("skinshowlog")) else Response.Write("")%></textarea>
        </strong></td>
    </tr>
    <tr>
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" �����޸� " >
      </div></td>
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
set rs=nothing
end sub

sub addskin()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">����û�ģ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
  <tr class="tdbg">
    <td height="30"><div align="center"><a href="admin_teamskin.asp?action=showskin"><strong>���ع���˵�</strong></a>���� <a href="admin_skin_help.asp" target="_blank"><strong>ģ���ǰ���</strong></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</div></td>
  </tr>
</table>
<form method="POST" action="admin_teamskin.asp" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>ģ�����</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">ģ�����ƣ�
        <input name="userskinname" type="text" id="userskinname">
        �������ߣ�
        <input name="skinauthor" type="text" id="skinauthor">
        �趨����:
        <select size=1 name="classid">
      	  <option value="0">------��δ����------</option>
          <%=sClasses%>
        </select>
        <br>
        ��������<strong>��
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="40" value="">
        </strong>
        Ԥ��ͼƬ<strong>��
        <input name="skinpic" type="text" id="skinpic" size="40">
        </strong> </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <strong>��ģ�壺</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"></textarea>
        <br>
        <br>
        <strong>��ģ�壺 <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"></textarea>
        </strong></td>
    </tr>
    <tr>
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveaddskin">
          <input name="cmdadd" type="submit" id="cmdadd" value=" ��� " >
      </div></td>
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

Sub SaveMove()
  dim id,ClassId
	id=Trim(Request("checkbox"))
	ClassId=Int(Trim(Request("classid")))
	id=FilterIds(id)
	If id<>"" Then
		oblog.execute("Update oblog_teamskin  Set classid= " & ClassId & " where id in ("&id&")")
	End If
	ReCountSkins
	oblog.ShowMsg "ģ��ת�Ƴɹ�",""
End Sub

Sub ReCountSkins()
	Dim rst,rst1
	Set rst=Server.CreateObject("Adodb.Recordset")
	Set rst1=Server.CreateObject("Adodb.Recordset")
	'���¼���
	rst.Open "select classid From oblog_skinclass WHERE itype = 1",conn,1,3
	rst1.Open "select Count(id) ,Classid From oblog_teamskin Where ispass=1 Group By classid",conn,1,3
	Do While Not rst.Eof
		rst1.Filter="classid=" & rst(0)
		If Not rst1.Eof Then
			oblog.Execute "Update oblog_skinclass Set icount=" & rst1(0) & " Where  itype = 1 AND classid=" & rst(0)
		Else
			oblog.Execute "Update oblog_skinclass Set icount=0 Where itype = 1 AND classid=" & rst(0)
		End If
		rst.MoveNext
	Loop
	oblog.execute "Update oblog_teamskin Set classid=0 Where classid Not In (select classid from oblog_skinclass Where itype = 1)"
	Set rst=Nothing
	Set rst1=Nothing
End Sub

Sub ShowClasses()
%>
<script language="javascript">
	function checkClass(){
		if(document.formC_0.classname1.value==""){
			alert("����д��������!");
			document.formC_0.classname1.focus();
			return false;
			}
			return true;
		}
</script>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">ģ �� �� �� ά ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <form name="formC_0" method="post" action="admin_teamskin.asp?action=saveclass" onsubmit="return checkClass();">
    <tr class="topbg">
      <td height="25" colspan="7" >
      	<strong>��������:<input type="text" size=20 name="classname1" maxlength=20></strong>
      	<input type="submit" value="����">
      </td>
    </tr>
  </form>
  <%If Not rstClass.Eof Then%>
    <tr class="topbg">
      <td width="20%" height="25" > <div align="center">������</div></td>
      <td width="40%" ><div align="center">��������</div></td>
      <td width="15%" > <div align="center">ģ����Ŀ</div></td>
      <td width="25%" ><div align="center">����</div></td>
    </tr>
    <%Do While Not rstClass.Eof %>
    <form id="formC_<%=rstClass("classid")%>" name="formC_<%=rstClass("classid")%>" method="post" action="admin_teamskin.asp?action=saveclass&classid=<%=rstClass("classid")%>">
    <tr class="tdbg">
      <td width="20%" height="25" > <div align="center"><%=rstClass("classid")%></div></td>
      <td width="40%" ><div align="center"><input type="text" name="classname1" value="<%=rstClass("classname")%>"></div></td>
      <td width="15%" > <div align="center"><%=rstClass("icount")%></div></td>
      <td width="25%" ><div align="center">
      	<input type="submit" value="�޸�"></a>&nbsp;&nbsp;|&nbsp;&nbsp;
      	<%
      	If rstClass("icount")>0 Then
      		%>
      		<input type="button" value="ɾ��" disabled>
      		<%
      	Else
      		%>
      		<input type="button" value="ɾ��" onclick="if(confirm('ȷ��Ҫɾ���÷�����?')==true) document.location.href='admin_teamskin.asp?action=delclass&classid=<%=rstClass("classid")%>'">
      		<%
      	End If
      	%>
      	</div></td>
    </tr>
  </form>
<%
		rstClass.Movenext
	Loop
%>
<%End if%>
 </table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%
End Sub
Sub SaveClass()
	Dim classid,classname,rst
	classname=ProtectSql(Left(Trim(Request("classname1")),20))
	classid=Request("classid")
	If classid<>"" Then classid=Int(classid)
	If classname<>"" Then
		Set rst=Server.CreateObject("ADODB.Recordset")
		rst.Open "select * From oblog_skinclass Where itype = 1 AND classname='" & classname & "'",conn,1,3
		'�Ƿ�����
		If Not rst.Eof Then
			rst.Close
			Set rst=Nothing
			'oblog.ShowMsg "Ŀ�����������з��������ظ�","admin_teamskin.asp?action=skinclass"
			Response.Redirect "admin_teamskin.asp?action=skinclass"
			Exit Sub
		End If

		If classid="" Then
			rst.AddNew
		Else
			rst.Close
			rst.Open "select * From oblog_skinclass Where itype = 1 AND classid=" & classid ,conn,1,3
		End If
		rst("classname")=classname
		If classid="" Then
			rst("icount")=0
			rst("itype")=1
		End If
		rst.update
		rst.Close
		Set rst=Nothing
		oblog.ShowMsg "��������ɹ�",""
	End If
End Sub
%>

<%
Sub DelClass()
	Dim classid,rst
	classid=Request("classid")
	If classid="" Then Exit Sub
	If classid<>"" Then classid=Int(classid)
	Set rst=Server.CreateObject("ADODB.Recordset")
	rst.Open "select * From oblog_skinclass Where itype = 1 AND classid=" & classid,conn,1,3
	If rst.Eof Then
		rst.Close
		Set rst=Nothing
		Response.Write "admin_teamskin.asp?action=skinclass"
	End If
	If rst("icount")>0 Then
		rst.Close
		Set rst=Nothing
		oblog.ShowMsg "Ŀ���������ģ������,�뽫�÷����е�ģ��ת�Ƶ���������Ȼ����ɾ��","admin_teamskin.asp?action=skinclass"
	Else
		rst.Delete
		Set rst=Nothing
		oblog.execute("update oblog_teamskin Set classid=0 where classid=" & classid)
		oblog.ShowMsg "�����ɾ���ɹ�","admin_teamskin.asp?action=skinclass"
	End If
End Sub
Set oblog = Nothing
%>
