<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
If CheckAccess("r_user_msg")=False Then Response.Write "��Ȩ����":Response.End
Dim rs, sql
Dim id, cmd, Keyword, sField
Keyword = Trim(Request("keyword"))
If Keyword <> "" Then Keyword = oblog.filt_badstr(Keyword)
sField = Trim(Request("Field"))
cmd = Trim(Request("cmd"))
Action = Trim(Request("Action"))
id = Trim(Request("id"))
If cmd = "" Then
    cmd = 0
Else
    cmd = CLng(cmd)
End If
G_P_FileName = "m_messages.asp?cmd=" & cmd & "&Field=" & sField & "&keyword=" & Keyword

%>
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
		<li class="main_top_left left">�� �� �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_messages.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>���ٲ��ң�</strong></td>
      <td width="687" height="30">
        <select name="Field" id="Field">
            <option value="author">����������</option>
            <option value="ip">������ip</option>
            <option value="userid">�û�ID</option>
            <option value="topic">���Ա���</option>
            <option value="content">��������</option>
        </select>
      <input type="hidden" name="cmd" value="2">
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit"  value=" ���� ">&nbsp;&nbsp;&nbsp;&nbsp;<a href="m_messages.asp">��������</a>|&nbsp;&nbsp;&nbsp;&nbsp;<a href="m_messages.asp?cmd=1">���������</a></td>
    </tr>
  </form>
  <form action="m_messages.asp" name=form2 method="get">
  <tr class="tdbg">
      <td width="100"><strong>��������</strong></td>
    <td>
            ��IP��������&nbsp;
            <input name="ip" type="text" id="ip" size="20" maxlength="30">
            <input type="checkbox"  name="chkIp" value="1" checked>�Ƿ񽫸�IP���뵽������
            <input type="hidden" name="action" value="clearip">
          <input type="submit"  value="����" />
        </td>
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
select Case Action
	Case "batchopt"
	    Call OptMessage("1")
	Case "clearip"
		Call OptMessage("2")
	Case Else
	    Call main
End select
If ErrMsg<>"" Then
    Call WriteErrMsg
End If

Sub Main()
    sql = "select top 500 userid,messagetopic,addtime,messageid,message_user,addip,message,iState,isguest From oblog_message "
    select Case cmd
        Case 0,""
        	sql= Sql & " Order By messageid desc"
            sGuide = sGuide & "��������"
        Case 1
        	sql= Sql & " Where iState=0 Order By messageid desc"
        	sGuide = sGuide & "���������"
        Case 2
            If Keyword = "" Then
            	ErrMsg="���󣺹ؼ��ֲ���Ϊ�գ�"
                Exit Sub
            Else
                select Case sField
	                Case "author"
	                    sql= Sql & " Where message_user like '%" & Keyword&"%' order by messageid desc"
	                    sGuide = sGuide & "�����������л��к���<font color=red> " & Keyword & " </font>������"
	                Case "userid"
	                    sql= Sql & " Where userid =" & Int(Keyword)&" order by messageid desc"
	                    sGuide = sGuide & "��������IDΪ<font color=red> " & Keyword & " </font>���ܵ�������"
	                Case "topic"
	                    sql= Sql & " Where messagetopic like '%" & Keyword & "%' order by messageid desc"
	                    sGuide = sGuide & "�����к��С� <font color=red>" & Keyword & "</font> ��������"
	                Case "ip"
	                    Sql= Sql & " Where addip='" & Keyword&"' order by messageid desc"
	                    sGuide = sGuide & "����ipΪ<font color=red> " & Keyword & " </font>������"
	                Case "content"
	                    sql= Sql & " Where message like '%" & Keyword&"%' order by messageid desc"
	                    sGuide = sGuide & "���������а���<font color=red> " & Keyword & " </font>������"
                End select
            End If
        Case Else
        	'Exit sub
    End select
    'Response.Write Sql
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, conn, 1, 1
    Call oblog.MakePageBar(rs, "ƪ����")
    rs.Close
    Set rs = Nothing
End Sub
Sub showContent()
    Dim i
    i = 0
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�� �� �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="myform" id="myform" method="post" action="m_messages.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
<style type="text/css">
<!--
td {padding:3px 0!important;}
-->
</style>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="0" style="word-wrap: break-word; word-break: break-all;">
          <%do while not rs.EOF %>
  <tr>
    <td align="center" style="background:#B3D1EA;border-bottom:1px #000 dotted;"  width="30"><input type="checkbox" name="chkOne" id="id" value='<%=rs("messageid")%>'></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;"><a href="../go.asp?messageid=<%=rs("messageid")%>" target="_blank" style="margin:0 0 0 10px;color:#333;"><%=oblog.filt_html(Left(rs("messagetopic"),20))%></a>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;color:#666;" width="300"><font color=#0d4d89><%=rs("message_user")%></font><%If rs("isguest") = 1 Then Response.Write "(�ο�)"%>&nbsp;������&nbsp;<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;"><%=rs("addtime")%></span>	��<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;">IP:<%=rs("addip")%></span></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="110">Ŀ���û�ID:<a href="../go.asp?userid=<%=rs("userid")%>" target="_blank" style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("userid")%></a></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="40">&nbsp;
<%If rs("iState")="1" Then %><span style="font-weight:600;color:#090;">����</span><%Else%><span style="font-weight:600;color:#f30;">����</span><%End If%>
</td>
  </tr>
  <tr>
    <td align="center" valign="top"><span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("messageid")%></span></td>
    <td colspan="4" valign="top" style="word-wrap: break-word; word-break: break-all;color:#555;"><%=Left(RemoveUBB(RemoveHtml(rs("message"))),200)%></td>
  </tr>
  <tr>
    <td height="8"></td>
    <td colspan="4"></td>
  </tr>
          <%
            i = i + 1
            If i >= G_P_PerMax Then Exit Do
            rs.MoveNext
        Loop
%>
</table>
 <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title">
            <td colspan=2 height=25>
                <input type="checkbox" name="chkAll" id="chkAll" onClick="javascript:CheckAll(this.form);">ȫѡ
                &nbsp;&nbsp;&nbsp;&nbsp;������ʽ:
				<input type="radio" name="opt" value="3">ɾ��
				<input type="radio" name="opt" value="2">ȡ�����
                <input type="radio" name="opt" value="1">ͨ�����&nbsp;&nbsp;
                <input type="hidden" value="batchopt" name="action">
                <input type="submit" value="��ʼ����" name="submit">
            </td>
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
End Sub

Sub OptMessage(sMod)
	Server.ScriptTimeOut=999999999
    Dim rstUser,rstCache,MSG
    'ɾ��/����/���
    Dim sOpt,sIds,sChkIp,sIP
    sIP=Request("ip")
	sChkIp=Request("chkip")
    sIds = FilterIds(Request("chkOne"))
    sOpt = Request("opt")
    '��ID����ɾ��
    If sMod="1" Then
		If Request.QueryString <> "" Then Exit Sub
	    If sOpt = "" Or sIds = "" Then Call main(): Exit Sub
		PassScore sIds,sOpt
	    select Case sOpt
	        Case "1"
				WriteSysLog "����������ͨ����˲�����Ŀ������ID��"&sIds&"",""
	            sql = "Update oblog_message Set iState=1 Where messageId In (" & sIds & ")"
	        Case "2"
				WriteSysLog "����������ȡ����˲�����Ŀ������ID��"&sIds&"",""
	            sql = "Update oblog_message Set iState=0 Where messageId In (" & sIds & ")"
	        Case "3"
				WriteSysLog "����������ɾ��������Ŀ������ID��"&sIds&"",""
	            sql = "Delete From oblog_message Where messageId In (" & sIds & ")"
	        Case Else
	            rstUser.Close
	            Set rstUser = Nothing
	            Exit Sub
	    End select
	    Set rstUser = oblog.Execute("select userid,count(messageid)  From oblog_message Where isdel=0 AND  istate=1 AND messageId In (" & sIds & ") Group By userid")
	    oblog.Execute sql
		MSG = "���Թ�������ɹ�!"
	Else
		MSG = "�������Գɹ�"
		sIp=oblog.filt_badstr(sIp)
		'OB_Debug Request("ip"),1
		If sIp="" Then Exit Sub
		Set rstUser = oblog.Execute("select userid,count(messageid)  From oblog_message Where isdel=0 AND istate=1 AND addip='" & sIP & "' Group By userid")
		oblog.Execute ("Delete From oblog_message Where addIp='" & sIp & "'")
		If sChkIp = "1" And oblog.ChkWhiteIP(sIP) = False Then
			'���������
			oblog.KillIP(sIP)
	    End If
		WriteSysLog "�������������������Ŀ������IP��"&sIp&"",oblog.NowUrl&"?"&Request.QueryString
	End If
    Dim blog
    Set blog = New Class_blog
    '���ⲿ���û������Խ��и���
    Do While Not rstUser.EOF
        '�����û����ּ�������Ŀ
        If sOpt = "3" Then
            sql = "update oblog_user set message_count=" & rstUser(1) & ",scores=scores-" & oblog.CacheScores(5)*rstUser(1) & " where userid=" & rstUser(0)
        Else
            sql = "update oblog_user set message_count=" & rstUser(1) & " where userid=" & rstUser(0)
        End If
        oblog.execute sql
        '���¾�̬ҳ��
        blog.userid = rstUser(0)
        blog.update_message 0
        blog.update_newmessage rstUser(0)
        rstUser.MoveNext
    Loop
    rstUser.Close
    Set rstUser = Nothing
    Set blog = Nothing
    oblog.ShowMsg MSG, ""
End Sub

Sub PassScore(id,iState)
	Dim rs,i
	Dim tid,sScore
	tid=id
	If iState= 1 Then
		sScore=oblog.CacheScores(5)
	Else
		sScore=-1*Abs(oblog.CacheScores(5))
	End if
	If InStr(tid,",")<0 Then
		Set rs = oblog.Execute ("select userid,istate FROM oblog_message WHERE messageid = " &tid)
		'����ǹ���
		If iState=1 Then
			'ֻ��������
			If rs(1)=0 Then
				oblog.GiveScore "",sScore,rs(0)
				oblog.execute("update oblog_user set message_count=message_count+1 where userid="&rs(0))
				oblog.execute("update oblog_setup set message_count=message_count+1")
			End if
		'�����ȡ����˻���ɾ��
		Else
			'ֻ�����Ѿ������
			If rs(1)=1 Then
				oblog.GiveScore "",sScore,rs(0)
				oblog.execute("update oblog_user set message_count=message_count-1 where userid="&rs(0))
				oblog.execute("update oblog_setup set message_count=message_count-1")
			End if
		End If
		rs.close
	Else
		tid = Split (tid ,",")
		For i = 0 To UBound(tid)
			Set rs = oblog.Execute ("select userid,istate FROM oblog_message WHERE messageid = " &tid(i))
			'����ǹ���
			If iState=1 Then
				'ֻ��������
				If rs(1)=0 Then
					oblog.GiveScore "",sScore,rs(0)
					oblog.execute("update oblog_user set message_count=message_count+1 where userid="&rs(0))
					oblog.execute("update oblog_setup set message_count=message_count+1")
				End if
			'�����ȡ����˻���ɾ��
			Else
				'ֻ�����Ѿ������
				If rs(1)=1 Then
					oblog.GiveScore "",sScore,rs(0)
					oblog.execute("update oblog_user set message_count=message_count-1 where userid="&rs(0))
					oblog.execute("update oblog_setup set message_count=message_count-1")
				End if
			End If
			rs.close
		Next
	End if
End Sub
Set oblog = Nothing
%>