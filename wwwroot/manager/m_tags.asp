<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
If CheckAccess("r_user_tag")=False Then Response.Write "��Ȩ����":Response.End
dim rs, sql
dim TagID,TagSearch,Keyword,strField

keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
strField=Trim(Request("Field"))
TagSearch=Trim(Request("TagSearch"))
Action=Trim(Request("Action"))
TagID=Trim(Request("TagID"))
'ComeUrl=Request.ServerVariables("HTTP_REFERER")
G_P_PerMax=20

if TagSearch="" then
	TagSearch=10
else
	TagSearch=CLng(TagSearch)
end if
G_P_FileName="m_tags.asp?TagSearch=" & TagSearch
if strField<>"" then
	G_P_FileName=G_P_FileName&"&Field="&strField
end if
if keyword<>"" then
	G_P_FileName=G_P_FileName&"&keyword="&keyword
end if
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
<title>oBlog--ϵ ͳ TAG �� ��</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">ϵ ͳ TAG �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_tags.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>���ٲ���TAG��</strong></td>
      <td  height="30"><select size=1 name="TagSearch" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="1">ʹ��Ƶ����ߵ�100��TAG</option>
          <option value="2">ʹ��Ƶ����͵�100��TAG</option>
          <option value="3">ʹ����Ϊ0��TAG</option>
          <option value="4">��������TAG</option>
        </select>       </td>
    </tr>
  </form>
  <form name="form2" method="post" action="m_tags.asp">
  <tr class="tdbg">
    <td width="100"><strong>�߼���ѯ��</strong></td>
    <td>
      <select name="Field" id="Field">
	  <option value="TagName" selected>TAG����</option>
      <option value="TagID" >TAG ID</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="TagSearch" type="hidden" id="TagSearch" value="10">
	  ��Ϊ�գ����ѯ����TAG</td>
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
if Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="merge" then
	call MergeTags()
elseif Action="batchlock" then
	call batchlock()
elseif Action="batchunlock" then
	call batchUnlock()
elseif Action="batchdel" Then
	call BatchDel
else
	call main()
end if
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	sGuide=""
	select case TagSearch
		case 1
			sql="select Top 100 * From oblog_tags Where iState=1 And iNum>0 order by iNum Desc"
			sGuide=sGuide & "ʹ��Ƶ����ߵ�100��TAG"
		case 2
			sql="select Top 100 * From oblog_tags Where iState=1 And iNum>0 order by iNum"
			sGuide=sGuide & "ʹ��Ƶ����͵�100��TAG"
		case 3
			sql="select  * From oblog_tags Where iState=1 And iNum=0"
			sGuide=sGuide & "ʹ����Ϊ0��TAG"
		case 4
			sql="select  * From oblog_tags Where iState=0"
			sGuide=sGuide & "��������TAG"
		case 10
			if Keyword="" then
				sql="select Top 100 * From oblog_tags Where  iNum>0 order by iNum Desc"
				sGuide=sGuide & "����TAG"
			else
				select case UCASE(strField)
				case "TAGID"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>TAG ID������������</li>"
					else
						sql="select * from oblog_tags where Tagid =" & CLng(Keyword)
						sGuide=sGuide & "TAG ID����<font color=red> " & CLng(Keyword) & " </font>"
					end if
				case "TAGNAME"
						sql="select * from oblog_tags where name like '%" & Keyword & "%' order by iNum Desc"
						sGuide=sGuide & "���С� <font color=red>" & Keyword & "</font> ����TAG"
				end select
			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>����Ĳ�����</li>"
	end select

	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set  rs=Server.CreateObject("Adodb.RecordSet")
	'Response.Write sql
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		sGuide=sGuide & "(<font color=red>0</font> )"
		%>
		<div id="main_body">
			<ul class="main_top">
				<li class="main_top_left left"><%=sGuide%></li>
				<li class="main_top_right right"> </li>
			</ul>
		</div>
		<%
	else
    	G_P_AllRecords=rs.recordcount
		sGuide=sGuide & "(<font color=red>" & G_P_AllRecords & "</font>)"
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
        	Response.write oblog.showpage(true,true,"�� TAG ")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rs.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	Response.write oblog.showpage(true,true,"�� TAG ")
        	else
	        	G_P_This=1
           		showContent
           		Response.write oblog.showpage(true,true,"�� TAG ")
	    	end if
		end if
	end if
	rs.Close
	Set  rs=Nothing
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

  <form name="myform" method="Post" action="m_tags.asp" onsubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title">
            <td width="30" align="center"><strong>ѡ��</strong></td>
            <td width="30" align="center"><strong>ID</strong></td>
            <td align="center"><strong>TAG����</strong></td>
            <td width="60" align="center"><strong>ʹ�ô���</strong></td>
            <td width="80" align="center"><strong>״̬</strong></td>
            <td width="60" align="center"><strong>����</strong></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg">
            <td align="center"><input name='TagID' type='checkbox' onclick="unselectall()" id="TagID" value='<%=cstr(rs("TagID"))%>'></td>
            <td  align="center"><span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("TagID")%></span></td>
            <td  align="Left" style="word-break: break-all; word-wrap:break-word;">&nbsp;&nbsp;<%
			Response.write "<a href='m_tags.asp?Action=Modify&TagID=" & rs("TagID") & "'"
			Response.write """>" & rs("Name") & "</a>"
			%> </td>
            <td  align="center">
			<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;">
			<%
			if rs("iNum")<>"" then
				Response.write rs("iNum")
			else
				Response.write "0"
			end if
			%>
			</span>
			</td>
            <td  align="center"><%
	  if rs("iState")=1 then
	  	Response.write "<span style=""font-weight:600;color:#090;"">����ʹ��</span>"
	  else
	  	Response.write "<span style=""font-weight:600;color:#f30;"">������</span>"
	  end if
	  %></td>
   <td  align="center"><%
		Response.write "<a href='m_tags.asp?Action=Modify&TagID=" & rs("TagID") & "'>�޸�</a>&nbsp;"
		if rs("iState")=1 then
			Response.write "<a href='m_tags.asp?Action=batchlock&TagID=" & rs("TagID") & "'>����</a>&nbsp;"
		else
      Response.write "<a href='m_tags.asp?Action=batchunlock&TagID=" & rs("TagID") & "'>����</a>&nbsp;"
		end if
		Response.write "<a href='m_tags.asp?Action=batchdel&TagID=" & rs("TagID") & "'>ɾ��</a>&nbsp;"
		%> </td>
          </tr>
          <%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
%>
        </table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ѡ�б�ҳ��ʾ������TAG</td>
            <td> <strong>������</strong>
            <%
              If TagSearch="4" Then
              %>
              <input name="Action" type="radio" value="batchunlock" checked onClick="document.myform.tagNames.disabled=true;document.myform.tagIds.disabled=true">
              �������
              <%
            Else
              %>
               <input name="Action" type="radio" value="batchlock" checked onClick="document.myform.tagNames.disabled=true;document.myform.tagIds.disabled=true">����
           <%
       		 End If
           %><input name="Action" type="radio" value="merge" onClick="document.myform.tagNames.disabled=false;document.myform.tagIds.disabled=false">�ϲ�Ϊ<input type="text" name="tagNames" id="tagNames" disabled>&nbsp;&nbsp;�ϲ����ID:<input type="text" name="tagIds" id="tagIds" size=10 disabled>
            <input name="Action" type="radio" value="batchdel"  onClick="document.myform.tagNames.disabled=true;document.myform.tagIds.disabled=true">ɾ��
              &nbsp;<input type="submit" name="Submit" value="ִ ��"></td>
  </tr>
</table></form>
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
	dim TagID
	dim rst,sSql
	TagID=Trim(Request("TagID"))
	if TagID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		TagID=CLng(TagID)
	end if
	Set  rst=Server.CreateObject("Adodb.RecordSet")
	sSql="select * from oblog_Tags where TagID=" & TagID
	if not IsObject(conn) then link_database
	rst.Open sSql,Conn,1,3
	if rst.bof and rst.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���� TAG ��</li>"
		rst.close
		Set  rst=nothing
		exit sub
	end if
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�޸�ע�� TAG ��Ϣ</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<FORM name="Form1" action="m_tags.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class="tdbg" >
      <TD width="40%"> TAG ����</TD>
      <TD width="60%"><input type="text" name="name" value="<%=rst("Name")%>" size=50></TD>
      <input type="hidden" value="<%=rst("Tagid")%>"  name="TagID">
    </TR>

    <TR class="tdbg" >
      <TD width="40%"> TAG ״̬��</TD>
      <TD width="60%"><input type="radio" name="iState" value=1 <%if rst("iState")=1 then Response.write "checked"%>>
        ����&nbsp;&nbsp; <input type="radio" name="iState" value=0 <%if rst("iState")=0 then Response.write "checked"%>>
        ����</TD>
    </TR>
    <TR class="tdbg" >
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"></TD>
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
	rst.close
	Set  rst=nothing
end sub

Sub BatchLock()
	Dim sID
	sId=FilterIds(Request("TagId"))
	If sId<>"" Then
		conn.Execute("Update oblog_Tags Set iState=0 Where TagId In (" & sID & ")")
		WriteSysLog "����������TAG������Ŀ��TAGID��"&sId&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "�����ɹ�!",""
	Else
		oblog.ShowMsg "��ѡ��Ҫ������TAG!",""
	End If
End Sub

Sub BatchUnLock()
	Dim sID
	sId=FilterIds(Request("TagId"))
	conn.Execute("Update oblog_Tags Set iState=1 Where TagId In (" & sID & ")")
	WriteSysLog "�����˽���TAG������Ŀ��TAGID��"&sId&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "�����ɹ�!",""
End Sub

Sub BatchDel()
	Dim sIDs,aIds,rst1,rst2,sTagIds1,sTags1,blog,sUserId
	sIDs=FilterIds(Request("TagId"))
	If sIds=""  Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	End If
	Set blog=New Class_blog
	'����ȡ��ʹ����ЩTag����־
	'��������TAG��������־
	'ɾ����ЩTAG
	Set rst1=Oblog.Execute("Select Distinct logid From oblog_usertags Where tagid in (" & sIds & ")")
	Do While Not rst1.Eof
		sTagIds1=""
		sTags1=""
		Set rst2=oblog.Execute("Select a.tagid,a.name,b.userid From oblog_tags a,oblog_usertags b Where a.tagid=b.tagid And  b.logid=" & rst1(0) & " And b.tagid Not in (" & sIds &")")
		If Not rst2.Eof Then
			Do While Not rst2.Eof
				sTagIds1=sTagIds1 & P_TAGS_SPLIT & rst2(0)
				sTags1=sTags1 & P_TAGS_SPLIT & rst2(1)
				sUserId=rst2(2)
				rst2.Movenext
			Loop
			If sTags1<>"" Then
				sTagIds1=Right(sTagIds1,Len(sTagIds1)-Len(P_TAGS_SPLIT))
				sTags1=Right(sTags1,Len(sTags1)-Len(P_TAGS_SPLIT))
			End If
			Call oblog.Execute("Update oblog_log Set logtags='" & sTags1 & "',logtagsid='" & sTagIds1 & "' Where logid=" & rst1(0))
			'���¾�̬�ļ�
			blog.userid = sUserId
	    blog.update_log rst1(0), 0
	  End If
		rst1.MoveNext
	Loop
	Set rst1=Nothing
	Set rst2=Nothing
	Set blog=Nothing
	conn.Execute("Delete From oblog_Tags Where TagId In (" & sIDs & ")")
	conn.Execute("Delete From oblog_UserTags Where TagId In (" & sIDs & ")")
	WriteSysLog "������ɾ��TAG������Ŀ��TAGID��"&sIds&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "ɾ���ɹ�!",""
End Sub

Sub MergeTags()
	If Request.QueryString <>"" Then Exit Sub
	Dim sIDs,sTargetId,sTargetName,aTags,i,sIDs0, rst,rst1,sSql,sTags,sTagsId,j
	sIDs=Trim(Request("TagId"))
	sTargetName=Trim(Request("tagNames"))
	sTargetId=Trim(Request("tagIds"))
	If sIds="" Or InStr(sIDs,",")=0 Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	End If
	sIDs=Replace(sIDs," ","")
	aTags=Split(sIDs,",")
	sIDs0=sIDs
	If Right(sIDs,1)<>"," Then sIDs=sIDs & ","
	If Left(sIDs,1)<>"," Then sIDs= "," & sIDs
	If Instr(1,sIDs,"," & sTargetId & ",",1)<=0 Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>Ŀ��ID����ȷ������Ϊ" & Replace(sIDs,","," ")&"֮�е�����һ������</li>"
		exit sub
	End If
	'���ȸ�oblog_Tags������
	Call conn.Execute("Update oblog_tags Set Name='" &sTargetName & "' Where TagId=" & sTargetId )
	'�滻����ID
	sIDs=Replace(sIDs,"," & sTargetId & ",","")
	If Left(sIDs,1)="," Then sIDs=Right(sIDs,Len(sIDs)-1)
	If Right(sIDs,1)="," Then sIDs=Left(sIDs,Len(sIDs)-1)
	'��ɾ����������,ע��SQL SERVER
	Call conn.Execute("Delete From  oblog_tags  Where TagId IN (" & sIDs & ")" )
	'�����Ѿ�ʹ�õ��û�TAG����¼��Щ��־��ID��Ȼ���������һ���µ�TAG ID����
	'��ȡΨһ����
	Set rst=Server.CreateObject("Adodb.Recordset")
	sSql="select a.logId,a.UserId From (select Userid,logId From oblog_Usertags Where TagId In (" & sIDs0 & ")) a Group by a.logId,a.UserId"
'	Response.Write sSql
	rst.Open sSql,conn,1,3
	'���½���ϵͳ����
	Call conn.Execute("Update oblog_tags Set iNum=" & rst.Recordcount & " Where TagId=" & sTargetId)
	'�����û�TAG������������
	Call conn.Execute("Delete From oblog_Usertags Where TagId In (" & sIDs0 & ")")
	'�������ݲ���
	Do While Not rst.Eof
		Call conn.Execute("Insert Into oblog_UserTags(tagid,userid,logid) Values(" & sTargetId &"," & rst("userid")& "," & rst("logid") & ")")
		'����������־���Tag
		Set rst1=conn.Execute("select b.* From oblog_UserTags a ,oblog_Tags b Where a.tagId=b.tagId And logid=" & rst("logid"))
		j=0
		sTags=""
		sTagsId=""
		'���TAG�ִ���ID�ִ�
		Do While Not rst1.Eof
			j=j+1
			If j=1 Then
				sTags=rst1("Name")
				sTagsId=rst1("TagId")
			Else
				sTags= sTags & P_TAGS_SPLIT & rst1("Name")
				sTagsId= sTagsId & P_TAGS_SPLIT & rst1("tagId")
			End if
			rst1.MoveNext
		Loop
		'���¹ؼ����ִ�
		Call conn.Execute("Update oblog_log Set logtags='" & sTags &"',logtagsid='" & sTagsId & "' Where logId=" & rst("logid"))
		'�������ɾ�̬ҳ�棿
		rst.Movenext
	Loop
	rst.close
	Set rst=Nothing
	Set rst1=Nothing
	WriteSysLog "�����˺ϲ�TAG������Ŀ��TAGID��"&sTargetId&"",""
	oblog.ShowMsg "TAG�ϲ��ɹ�!",""
End Sub

sub SaveModify()
	If Request.QueryString <>"" Then Exit Sub
	dim sID,sName,sState,rst,sSql
	sName=Trim(Request.Form("Name"))
	sID=Trim(Request.Form("TagID"))
	sState=Trim(Request.Form("iState"))
	if sID="" Or Not IsNumeric(sId) then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		sID=CLng(sID)
	end if

	if founderr=true then
		exit sub
	end if

	Set  rst=Server.CreateObject("Adodb.RecordSet")
	sSql="select * from oblog_Tags where TagID=" & sID
	if not IsObject(conn) then link_database
	rst.Open sSql,Conn,1,3

	rst("Name")=sName
	rst("iState")=sState
	rst.update
	rst.Close
	Set  rst=Nothing
	WriteSysLog "�������޸�TAG������Ŀ��TAGID��"&sID&"",""
	oblog.ShowMsg "�޸ĳɹ�!",""
end sub

sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charSet =gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	Response.write strErr
end sub
Set oblog = Nothing
%>