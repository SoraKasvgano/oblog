<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
If CheckAccess("r_user_blog")=False Then Response.Write "��Ȩ����":Response.End
dim rs, sql
dim id,cmd,Keyword,sField,sDate1,sDate2,sRsClass,sclass
Dim douname
'-----------------------------
Dim Z_logRole,Z_classRole
	If oblog.CheckAdmin(1) Then
		Z_classRole=" "
		Else
		Z_logRole=session("r_classes1")
'		OB_DEBUG Z_logRole,1
		If Len(z_logrole) > 0 Or Not IsNull(z_logrole) Then
			If InStr(z_logrole,",") Then
				Dim rsmain,ustr
				Dim strTemp,arrTemp,j
				arrTemp = Split (z_logrole, ",")
				For j = 0 To UBound(arrTemp)
					set rsmain=oblog.execute("select id from oblog_logclass where parentpath like '"&arrTemp(j)&",%' OR parentpath like '%,"&arrTemp(j)&"' OR parentpath like '%,"&arrTemp(j)&",%'")
					while not rsmain.eof
						ustr=ustr&","&rsmain(0)
						rsmain.movenext
					Wend
					If Left(ustr,1)="," Then
						ustr=arrTemp(j)&ustr
					Else
						ustr=arrTemp(j)& "," &ustr
					End If
				Next
				Z_classRole = FilterIDs(ustr)
'				OB_DEBUG Z_classRole,1
				Z_classRole=" and classid in("&Z_classRole&") "
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
Action=LCase(Trim(Request("Action")))
douname=lcase(trim(oblog.filt_badstr(request("douname"))))
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
G_P_FileName="m_blog.asp?cmd=" & cmd & "&field=" & sField & "&keyword=" & keyword & "&date1=" & sDate1 & "&date2=" &sDate2
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
		<li class="main_top_left left">�� ־ �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_blog.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>���ٲ��ң�</strong></td>
      <td width="687" height="30"><select size=1 name="cmd" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="0">����500ƪ��־</option>
          <option value="1">�������־</option>
          <option value="2">δͨ����˵���־</option>
          <option value="3">������־</option>
<!--           <option value="4">����˵ľ�����־</option> -->
          <option value="9">������־</option>
          <option value="10">������־</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="m_blog.asp">��־������ҳ</a>|&nbsp;&nbsp;<a href="m_blog.asp?cmd=9">�����б�</a>|&nbsp;&nbsp;<a href="m_blog.asp?cmd=3">�����б�</a></td>
    </tr>
  </form>
<!--    <form name="form2" action="m_user.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>����־�����ѯ��</strong></td>
      <td width="687" height="30">
      	<select size=1 name="classid">
      	  <option value="0">------ȫ������------</option>
          <%=sClass%>
        </select>
		 <input name="cmd" type="hidden" id="cmd" value="109">
        <input type="submit" value=" �� ѯ "></td>
    </tr>
  </form> -->
  <form name="form2" method="post" action="m_blog.asp">
  <tr class="tdbg">
      <td width="120"><strong>�߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
	      <option value="author" selected>�û�����</option>
		  <option value="logid" >��־ID</option>
	      <option value="userid" >�û�ID</option>
	      <option value="ip">����IP</option>
	      <option value="title" >��������</option>
	      <option value="content" >��������</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="cmd" type="hidden" id="cmd" value="10">
        ��Ϊ�գ����ѯ����</td>
  </tr>
</form>
  <form name="form3" method="post" action="m_blog.asp">
  <tr class="tdbg">
      <td width="120"><strong>��ʱ�����β�ѯ��</strong></td>
    <td>
    	��ʼʱ�䣺<input type="text" name="date1" size=14 maxlength=14>
    	����ʱ�䣺<input type="text" name="date2" size=14 maxlength=14>

      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="cmd" type="hidden" id="cmd" value="11">
      <br/>
        ʱ���ʽ��YYYYMMDDHHMm����2006��6��6��9��12�֣�������200606060912,������ʽ����֧��</td>
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
If action = "del" Or action = "best0" Or action = "best1" Or action = "pass0" Or action = "pass1" Or action = "move" Or action = "moveclass" Then
	If id = "" Then
		oblog.ShowMsg "������ѡ��һ��ID���в���" , ""
	End If
End If
select Case LCase(action)
	Case "modify"
		call Modify()
	Case "savemodify"
		call SaveModify()
	Case "del"
		Call DelScore(id)
		oblog.execute("update oblog_log Set isdel=1 where logid In ("&id & ")")
		'ɾ����־�ļ�!
		delblogs id
		WriteSysLog "ɾ���� "&douname&" ��һƪIDΪ("&id&")����־.���������վ��",oblog.NowUrl&"?"&OB_IIF(Request.QueryString,Request.Form)
		oblog.ShowMsg "ɾ���ɹ���",""
	Case "best0"
		Call BestScore(id,0)
		oblog.execute("update oblog_log Set isbest=0 Where logid In (" & id & ")")
'		Response.Redirect "m_blog.asp?cmd=3"
		WriteSysLog "ȡ���� "&douname&" ��һƪIDΪ("&id&")����־�ľ���.",oblog.NowUrl&"?"&OB_IIF(Request.QueryString,Request.Form)
		oblog.ShowMsg "ȡ�������ɹ���",""
	Case "best1"
		Call BestScore(id,1)
		oblog.execute("update oblog_log Set isbest=1 Where logid In (" & id & ")")
'		Response.Redirect "m_blog.asp?cmd=3"
		WriteSysLog "�� "&douname&" ��IDΪ("&id&")����Ϊ("&oblog.filt_badstr(unescape(request("title")))&")����־����Ϊ����",oblog.NowUrl&"?"&OB_IIF(Request.QueryString,Request.Form)
		If int(oblog.CacheConfig(86)) = 1 Then
		oblog.execute("INSERT INTO oblog_pm(incept,sender,topic,content) VALUES('"&doUname&"','ϵͳ����Ա','ϵͳ֪ͨ!�������±���Ϊ����!','��ϲ,���ı���Ϊ   "&id&".   "&oblog.filt_badstr(unescape(request("title")))&"      ����־,�Ѿ�������Ա��Ϊ����!�ٽ�����Ŷ!(����Ϣϵͳ�Զ�����,�Ķ��󽫱��Զ�ɾ��.�����ػظ�!)')")
		End If
		oblog.ShowMsg "���þ����ɹ���",""
	Case "pass0"
		oblog.execute("update oblog_log Set passcheck=0 Where logid In (" & id & ")")
		oblog.execute("update oblog_userdigg Set iState=0 Where logid In (" & id & ")")
		'������־����
		DoUpdatelog id
'		Response.Redirect "m_blog.asp"
		WriteSysLog "��������־ȡ����˲�����Ŀ����־ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "��������־Ϊδ���״̬��",""
	Case "pass1"
		oblog.execute("update oblog_log Set passcheck=1 Where logid In (" & id & ")")
		oblog.execute("update oblog_userdigg Set iState=1 Where logid In (" & id & ")")
		'������־����
		DoUpdatelog id
'		Response.Redirect "m_blog.asp"
		WriteSysLog "��������־ͨ����˲�����Ŀ����־ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "��������־Ϊ���״̬��",""
	Case "move"
		oblog.execute("update oblog_log Set specialid=" & clng(Request("SpecialId")) &" Where logid In (" & id & ")")
'		Response.Redirect "m_blog.asp"
		WriteSysLog "��������־ת�Ʋ�����Ŀ����־ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "��־ת�Ƴɹ���",""
	Case "moveclass"
		oblog.execute("update oblog_log Set classid=" & clng(Request("classid")) &" Where logid In (" & id & ")")
'		Response.Redirect "m_blog.asp"
		WriteSysLog "��������־����ת�Ʋ�����Ŀ����־ID��"&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "��־����ת�Ƴɹ���",""
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
		case 0
			sql="select " & sQryFields & " from oblog_log  Where isdel=0 " & Z_classRole & " order by logid desc"
			sGuide=sGuide & "����500ƪ��־"
		case 1
			sql="select " & sQryFields & " from oblog_log where passcheck=1 And isdel=0 " & Z_classRole & " order by logid desc"
			sGuide=sGuide & "ͨ����˵���־"
		case 2
			sql="select " & sQryFields & " from oblog_log where passcheck=0 and isdraft=0 And isdel=0 " & Z_classRole & " order by logid desc"
			sGuide=sGuide & "δͨ����˵���־"
		Case 3
			sql="select " & sQryFields & " from oblog_log where passcheck=1 And isdel=0  and isbest=1 " & Z_classRole & " order by logid desc"
			sGuide=sGuide & "������־"
		Case 4
'			sql="select " & sQryFields & " from oblog_log where passcheck=1 And isdel=0  and isbest=2 order by logid desc"
'			sGuide=sGuide & "����˵ľ�����־"
		Case 9
			sql="select " & sQryFields & " from oblog_log where isTrouble=1  And isdel=0 " & Z_classRole & " order by logid desc"
			sGuide=sGuide & "������־"
		case 10
			if Keyword="" then
				sql="select " & sQryFields & " from oblog_log  Where isdel=0  " & Z_classRole & " order by logid desc"
				sGuide=sGuide & "������־"
			else
				select case sField
				case "logid"
					if Not IsNumeric(Keyword) then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID������������</li>"
					else
						sql="select " & sQryFields & " from oblog_log where isdel=0  and logid =" & CLng(Keyword) & Z_classRole
						sGuide=sGuide & "��־ID����<font color=red> " & CLng(Keyword) & " </font>����־"
					end if
				case "userid"
					if Not IsNumeric(Keyword) then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID������������</li>"
					else
						sql="select " & sQryFields & " from oblog_log where isdel=0  and userid =" & CLng(Keyword) & Z_classRole
						sGuide=sGuide & "����ID����<font color=red> " & CLng(Keyword) & " </font>����־"
					end if
				case "author"
					sql="select " & sQryFields & " from oblog_log where  isdel=0  and author like '%" & Keyword & "%' " & Z_classRole & "  order by logid  desc"
					sGuide=sGuide & "���������к��С� <font color=red>" & Keyword & "</font> ������־"
				case "ip"
					sql="select " & sQryFields & " from oblog_log where  isdel=0  and addip like '%" & Keyword & "%'  " & Z_classRole & " order by logid  desc"
					sGuide=sGuide & "������־ʱ��IP�к��С� <font color=red>" & Keyword & "</font> ������־"
				case "title"
					sql="select " & sQryFields & " from oblog_log where  isdel=0  and topic like '%" & Keyword & "%'  " & Z_classRole & " order by logid  desc"
					sGuide=sGuide & "��־�����к��С� <font color=red>" & Keyword & "</font> ������־"
				case "content"
					sql="select " & sQryFields & " from oblog_log where  isdel=0  and logtext like '%" & Keyword & "%' " & Z_classRole & "  order by logid  desc"
					sGuide=sGuide & "��־�����к��С� <font color=red>" & Keyword & "</font> ������־"
				end select
			end if
		Case 11
			sDate1=DeDateCode(sDate1)
			sDate2=DeDateCode(sDate2)
			If sDate1<>"" And sDate2<>"" Then
				sql="select " & sQryFields & " from oblog_log where truetime>=" & G_Sql_d_Char & sDate1 & G_Sql_d_Char & " And  truetime<=" & G_Sql_d_Char & sDate2 & G_Sql_d_Char &  " and isdel=0  " & Z_classRole & " order by logid  desc"
				sGuide=sGuide & "ʵ�ʷ���ʱ���� <font color=red>" & sDate1 & "</font> �� <font color=red>" & sDate2 & "</font> ����־"
			End If
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>����Ĳ�����</li>"
	end Select
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
'	OB_DEBUG sql,1
	If Trim(Sql)="" Then
		oblog.ShowMsg "��������ȷ�Ĳ�ѯ������",""
	End If
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
  <form name="myform" method="Post" action="m_blog.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="0" style="word-break:break-all;">
          <%do while not rs.EOF %>
  <tr>
    <td align="center" style="background:#B3D1EA;border-bottom:1px #000 dotted;" width="30">
    	<input name='id' type='checkbox' onclick="unselectall()" id="id" value='<%=cstr(rs("logid"))%>'>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;"><span>[<%=oblog.GetClassName(2,0,rs("classid"))%>]</span>
    	<a href="../go.asp?logid=<%=rs("logid")%>" target="_blank" style="margin:0 0 0 10px;color:#333;"><%=oblog.Filt_html(rs("topic"))%></a>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="290"><a href="../go.asp?userid=<%=rs("userid")%>" target="_blank"><font color=#0d4d89><%=rs("author")%></font></a>&nbsp;������
	<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;">
	<%
		Response.write rs("addtime") & "</span>��<span style=""font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;color:#777;"">IP:" &  rs("addip")
	%></span>
	</td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="60">
		<%
			select case rs("passcheck")
				case 0
					Response.write "<span style=""font-weight:600;color:#f30;"">�ȴ����</span>"
				case 1
					Response.write "<span style=""font-weight:600;color:#090;"">ͨ�����</span>"
			end select
		%>
	</td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="108">
<%
        Response.write "<a href='../admin_edit.asp?do=5&Action=modilog&logid=" & rs("logid") & "'>�޸�</a>&nbsp;"
		If rs("isbest")=1 Then
        	Response.write "<a href='m_blog.asp?Action=best0&id=" & rs("logid") & "&douname="&rs("author")&"' onClick='return confirm(""ȷ��Ҫȡ������־�ľ���������"");'><font color=red>ȡ��</font></a>&nbsp;"
        Else
        	Response.write "<a href='m_blog.asp?Action=best1&id=" & rs("logid") & "&douname="&rs("author")&"&title="&escape(rs("topic"))&"' onClick='return confirm(""ȷ��Ҫ���ø���־Ϊ������"");'>�Ӿ�</a>&nbsp;"
        End  If

        Response.write "<a href='m_blog.asp?Action=Del&id=" & rs("logid") & "&douname="&rs("author")&"' onClick='return confirm(""ȷ��Ҫɾ������־��"");'>ɾ��</a>&nbsp;"
%>
</td>
  </tr>
  <tr>
    <td align="center" valign="top"><span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("logid")%></span></td>
    <td colspan="4" valign="top" style="word-wrap: break-word; word-break: break-all;color:#555;"><%=Left(oblog.Filt_html(RemoveHtml(rs("logtext"))),200)%></td>
  </tr>
  <tr>
    <td height="8" colspan="5"></td>
  </tr>
<%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
rs.Close
%>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="140" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ѡ�б�ҳ������־</td>
            <td> <strong>������</strong>
              <input name="Action" type="radio" value="Del">
              ɾ��&nbsp;&nbsp;
              <input name="Action" type="radio" value="pass0">
              ����&nbsp;&nbsp;
              <input name="Action" type="radio" value="pass1"">
              ���&nbsp;&nbsp;
              <input name="Action" type="radio" value="best1">
              ����&nbsp;&nbsp;
              <input name="Action" type="radio" value="best0">
              ȡ������&nbsp;&nbsp;
              <input name="Action" type="radio" value="moveclass" onClick="document.myform.classid.disabled=false">
              ת��&nbsp;&nbsp;
<!--               <input name="Action" type="radio" value="Move" onClick="document.myform.SpecialId.disabled=false">
              <select name="SpecialId" id="SpecialId" disabled>
              	<option value=0>ȡ��ר������</option>
								<%
								Set rs = oblog.Execute("select specialid,s_name From oblog_Special Where isActive=1 Order By SpecialId Desc")
								Do While Not rs.Eof
								%>
                <option value=<%=rs(0)%>><%=Left(rs(1),7)%></option>
                <%
	                rs.Movenext
	              Loop
	              Set rs=Nothing
                %>
              </select>
              &nbsp;&nbsp; -->
			<select name="classid" id="classid" disabled>
			<%=oblog.show_class("log",0,0)%>
			</select>
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
    Dim  rs, blog, i
    Set rs = oblog.execute("select userid,logid from oblog_log where logid in (" & ids & ")")
    Set blog = New class_blog
    Do While Not rs.Eof
        blog.userid = rs(0)
		blog.Update_index 0
        blog.update_log rs(1), 0
        rs.movenext
    Loop
    Set rs = Nothing
    Set blog = Nothing
End Sub

Sub delblogs(ids)
    Dim uid, delname, rst, fso, sid,i,logid,blog,cid
    Set fso = Server.CreateObject(oblog.CacheCompont(1))
    logid=Split(ids,",")
    Set rst = Server.CreateObject("adodb.recordset")
	Set blog = New class_blog
    For i=0 To UBound(logid)
	    rst.open "select a.userid,a.logfile,a.subjectid,a.logtype,a.scores,a.isdel,b.user_dir,b.user_folder,a.classid from oblog_log a ,oblog_user b where a.userid=b.userid And logid="&logid(i),conn,1,3
	    If Not rst.Eof Then
				uid = rst(0)
				delname = OB_IIF(Trim(rst(1)),"")
				sid = rst(2)
				cid = rst(8)
				'�����ļ���¼
				'Call oblog.DeleteFiles(logid)
				'��ʵ������Ҫ���������ļ�����
				'�����ļ���ʱɾ��
				'If true_domain = 1 And delname <> "" Then
				    If InStr(delname, "archives") Then
				        delname = Right(delname, Len(delname) - InStrRev(delname, "archives") + 1)
				    Else
				        delname = Right(delname, Len(delname) - InStrRev(delname, "/"))
				    End If
				    delname=blogdir & rst("user_dir")& "/" & rst("user_folder")&"/"&oblog.l_ufolder&"/"&delname
				'End If
				If delname <> "" Then
						delname=Replace(delname,"//","/")
				    If fso.FileExists(Server.MapPath(delname)) Then fso.DeleteFile Server.MapPath(delname)
				End If

				'--------------------------------------------
				'���¼�����,ɾ������
				If rst("isdel")=1 Then
					Call Tags_UserDelete(logid(i))
					Call OBLOG.log_count(uid,logid(i),sid,cid,"-")
				End If
				'--------------------------------------------
				rst.Close
				blog.userid = uid
				blog.Update_Subject uid
				blog.Update_index 0
				blog.Update_newblog (uid)
			End If
  	Next
	Set blog = Nothing
	Set fso = Nothing
	Set rst = Nothing
End Sub

Sub BestScore(id,iState)
	Dim rs,i
	Dim tid,sScore
	tid=id
	If iState= 1 Then
		sScore=oblog.CacheScores(9)
	Else
		sScore=-1*Abs(oblog.CacheScores(9))
	End If
	If InStr(tid,",")<0 Then
		Set rs = oblog.Execute ("select userid,isbest FROM oblog_log WHERE logid = " &tid)
		'ֻ����֮ǰδ�Ӿ���
		If iState= 1 Then
			If rs(1)=0 Then oblog.GiveScore "" ,sScore,rs(0)
		Else
			If rs(1)=1 Then oblog.GiveScore "" ,sScore,rs(0)
		End if
		rs.close
	Else
		tid = Split (tid ,",")
		For i = 0 To UBound(tid)
			Set rs= oblog.execute ("select userid,isbest FROM oblog_log WHERE logid = " &tid(i))
			If iState= 1 Then
				If rs(1)=0 Then oblog.GiveScore "" ,sScore,rs(0)
			Else
				If rs(1)=1 Then oblog.GiveScore "" ,sScore,rs(0)
			End if
			rs.close
		Next
	End if
End Sub

Sub DelScore(id)
	Dim rs,i
	Dim tid,sScore
	tid=id
	'ɾ����־ʱ����ɾ������־����õ����л���,���ҽ��л��ֳͷ�
	If InStr(tid,",")<0 Then
		Set rs = oblog.Execute ("select userid,scores FROM oblog_log WHERE logid = " &tid)
		sScore=-1*(rs(1)+CLng(oblog.CacheScores(4)))
		If IsNull(sScore) Then sScore = -1*(CLng(oblog.CacheScores(4)))
		oblog.GiveScore "",sScore,rs(0)
		rs.close
	Else
		tid = Split (tid ,",")
		For i = 0 To UBound(tid)
			Set rs= oblog.execute ("select userid,scores FROM oblog_log WHERE logid = " &tid(i))
			sScore=-1*(rs(1)+CLng(oblog.CacheScores(4)))
			If IsNull(sScore) Then sScore = -1*(CLng(oblog.CacheScores(4)))
			oblog.GiveScore "",sScore,rs(0)
			rs.close
		Next
	End if
End Sub
Set oblog = Nothing
%>