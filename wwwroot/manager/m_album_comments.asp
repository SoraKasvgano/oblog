<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_album_comment")=False Then Response.Write "��Ȩ����":Response.End
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
G_P_FileName = "m_album_comments.asp?cmd=" & cmd & "&Field=" & sField & "&keyword=" & Keyword

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
		<li class="main_top_left left">�� �� �� �� �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_album_comments.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>���ٲ��ң�</strong></td>
      <td width="687" height="30">
        <select name="Field" id="Field">
            <option value="author">����������</option>
            <option value="ip">������ip</option>
            <option value="userid">�û�ID</option>
            <option value="topic">���۱���</option>
            <option value="content">��������</option>
        </select>
      <input type="hidden" name="cmd" value="2">
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit"  value=" ���� ">&nbsp;&nbsp;&nbsp;&nbsp;<a href="m_album_comments.asp">��������</a>|&nbsp;&nbsp;&nbsp;&nbsp;<a href="m_album_comments.asp?cmd=1">���������</a></td>
    </tr>
  </form>
  <form action="m_album_comments.asp" name="form2" method="get">
  <tr class="tdbg">
      <td width="100"><strong>��������</strong></td>
    <td>
            ��IP��������&nbsp;
            <input name="ip" type="text" size="20" maxlength="30">
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
	    Call Optcomment("1")
	Case "clearip"
		Call Optcomment("2")
	Case Else
	    Call main
End select
If ErrMsg<>"" Then
    Call WriteErrMsg
End If

Sub Main()
    sql = "select top 500 userid,commenttopic,addtime,commentid,comment_user,addip,comment,iState,mainid,isguest From oblog_albumcomment "
    select Case cmd
        Case 0,""
        	sql= Sql & " Order By commentid desc"
            sGuide = sGuide & "��������"
        Case 1
        	sql= Sql & " Where iState=0 Order By commentid desc"
        	sGuide = sGuide & "���������"
        Case 2
            If Keyword = "" Then
            	ErrMsg="���󣺹ؼ��ֲ���Ϊ�գ�"
                Exit Sub
            Else
                select Case sField
	                Case "author"
	                    sql= Sql & " Where comment_user like '%" & Keyword&"%' order by commentid desc"
	                    sGuide = sGuide & "�����������л��к���<font color=red> " & Keyword & " </font>������"
	                Case "userid"
	                    sql= Sql & " Where userid =" & Int(Keyword)&" order by commentid desc"
	                    sGuide = sGuide & "��������IDΪ<font color=red> " & Keyword & " </font>���ܵ�������"
	                Case "topic"
	                    sql= Sql & " Where commenttopic like '%" & Keyword & "%' order by commentid desc"
	                    sGuide = sGuide & "�����к��С� <font color=red>" & Keyword & "</font> ��������"
	                Case "ip"
	                    Sql= Sql & " Where addip='" & Keyword&"' order by commentid desc"
	                    sGuide = sGuide & "����ipΪ<font color=red> " & Keyword & " </font>������"
	                Case "content"
	                    sql= Sql & " Where comment like '%" & Keyword&"%' order by commentid desc"
	                    sGuide = sGuide & "���������а���<font color=red> " & Keyword & " </font>������"
                End select
            End If
        Case Else
        	Exit sub
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
		<li class="main_top_left left"><%=sGuide%></li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="myform" id="myform" method="post" action="m_album_comments.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
<style type="text/css">
<!--
td {padding:3px 0!important;}
-->
</style>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="0" style="word-wrap: break-word; word-break: break-all;">
          <%do while not rs.EOF %>
  <tr>
    <td align="center" style="background:#B3D1EA;border-bottom:1px #000 dotted;" width="30"><input type="checkbox" name="chkOne" id="id" value='<%=rs("commentid")%>'></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;"><a href="../go.asp?fileid=<%=rs("mainid")%>#<%=rs("commentid")%>" target="_blank" style="margin:0 0 0 10px;color:#333;"><%=oblog.filt_html(RemoveHtml(Left(rs("commenttopic"),20)))%></a></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;color:#666;" width="300"><font color=#0d4d89><%=rs("comment_user")%></font><%If rs("isguest") = 1 Then Response.Write "(�ο�)"%>&nbsp;������&nbsp;<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;"><%=rs("addtime")%></span>	��<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;">IP:<%=rs("addip")%></span></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="110">Ŀ���û�ID:<a href="../go.asp?userid=<%=rs("userid")%>" target="_blank" style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("userid")%></a></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="40">&nbsp;
<%If rs("iState")="1" Then %><span style="font-weight:600;color:#090;">����</span><%Else%><span style="font-weight:600;color:#f30;">����</span><%End If%>
</td>
  </tr>
  <tr>
    <td align="center" valign="top"><span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("commentid")%></span></td>
    <td colspan="4" valign="top" style="word-wrap: break-word; word-break: break-all;"><%=Left(RemoveUBB(RemoveHtml(rs("comment"))),100) & "..."%></td>
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
                &nbsp;&nbsp;&nbsp;&nbsp;
                ������ʽ:
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

Sub Optcomment(sMod)
	Server.ScriptTimeOut=999999999
    Dim rstUser,rstCache,rstBlog,rstComment,MSG
    'ɾ��/����/���
    Dim sOpt,sIds,sChkIp,sIP
    sIP=Request("ip")
	sChkIp=Request("chkip")
    sIds = FilterIds(Request("chkOne"))
    sOpt = Request("opt")
    '��ID����ɾ��
    If sMod="1" Then
	    If sOpt = "" Or sIds = "" Then Call main(): Exit Sub
	    'ȡ��һuserid
	    'Set rstUser = oblog.Execute("select b.userid,count(a.commentid) From oblog_albumcomment a,(select userid From oblog_albumcomment Where commentId In (" & sIds & ") Group By userid) b Where a.userid=b.userid")
	    Set rstUser = oblog.Execute("select userid From oblog_albumcomment Where commentId In (" & sIds & ") Group By userid")
	    'ȡ��һ��logid
	    Set rstBlog = oblog.Execute("select mainid From oblog_albumcomment Where commentid In (" & sIds & ") Group By mainid")
		PassScore sIds,sOpt
	    select Case sOpt
	        Case "1"
	            '�ȸ��֣������Ի�ó�ʼ״̬
				WriteSysLog "�������������ͨ����˲�����Ŀ������ID��"&sIds&"",""
	            sql = "Update oblog_albumcomment Set iState=1 Where commentId In (" & sIds & ")"
	        Case "2"
				WriteSysLog "�������������ȡ����˲�����Ŀ������ID��"&sIds&"",""
	            sql = "Update oblog_albumcomment Set iState=0 Where commentId In (" & sIds & ")"
	        Case "3"
				WriteSysLog "�������������ɾ��������Ŀ������ID��"&sIds&"",""
	            sql = "Delete From oblog_albumcomment Where commentId In (" & sIds & ")"
	        Case Else
	            rstUser.Close
	            Set rstUser = Nothing
	            Exit Sub
	    End select
	    oblog.Execute sql
		MSG = "���۹�������ɹ�!"
	Else
		MSG = "�������۳ɹ�"
		sIp=oblog.filt_badstr(sIp)
		'OB_Debug Request("ip"),1
		If sIp="" Then Exit Sub
		'Set rstUser = oblog.Execute("select userid,count(commentid)  From oblog_albumcomment Where addip='" & sIP & "' Group By userid")
		Set rstUser = oblog.Execute("select userid From oblog_albumcomment Where addip='" & sIP & "' Group By userid")
		Set rstBlog = oblog.Execute("select mainid From oblog_albumcomment Where addip='" & sIP & "' Group By mainid")
		oblog.Execute ("Delete From oblog_albumcomment Where addIp='" & sIp & "'")
		If sChkIp = "1" And oblog.ChkWhiteIP(sIP) = False Then
			'���������
			oblog.KillIP(sIP)
	    End If
		WriteSysLog "����������������������Ŀ������IP��"&sIp&"",oblog.NowUrl&"?"&Request.QueryString
	End If
    Dim blogcomments,allComments
    '����־�������½��м���
    Do While Not rstUser.EOF
        '�����û�����
        Set rstComment=oblog.Execute("select Count(commentid) From oblog_comment Where istate=1 AND userid=" & rstUser(0))
		allComments =  rstComment(0)
		rstComment.Close
        Set rstComment=oblog.Execute("select Count(commentid) From oblog_albumcomment Where istate=1 AND userid=" & rstUser(0))
		allComments = allComments + rstComment(0)
        '������Ŀ
        If sOpt = "3" Then
            sql = "update oblog_user set comment_count=" & allComments & ",scores=scores-" & oblog.CacheScores(6)*rstComment(0) & " where userid=" & rstUser(0)
        Else
            sql = "update oblog_user set comment_count=" & allComments & " where userid=" & rstUser(0)
        End If
        oblog.Execute Sql
        rstUser.MoveNext
    Loop
    Set rstComment=Nothing
    rstUser.Close
    Do While Not rstBlog.Eof
    	Set rstUser=oblog.Execute("select count(commentid) From oblog_albumcomment Where istate=1 AND mainid=" & rstBlog(0))
    	If rstUser.Eof Then
    		blogcomments=0
    	Else
    		blogcomments=rstUser(0)
    	End If
        '���¼���������Ŀ
        oblog.Execute ("update [oblog_album] set commentnum=" & blogcomments  & " Where fileid=" & rstBlog(0))
        rstBlog.MoveNext
    Loop
    rstBlog.Close
    Set rstUser = Nothing
    Set rstBlog = Nothing
    oblog.ShowMsg MSG, ""
End Sub
'iState=1 ͨ�����;2ȡ�����;3ɾ��
Sub PassScore(id,iState)
	Dim rs,i
	Dim tid,sScore
	tid=id
	If iState= 1 Then
		sScore=oblog.CacheScores(6)
	Else
		sScore=-1*Abs(oblog.CacheScores(6))
	End if
	If InStr(tid,",")<0 Then
		Set rs = oblog.Execute ("select userid,istate FROM oblog_albumcomment WHERE commentid = " &tid)
		'����ǹ���
		If iState=1 Then
			'ֻ��������
			If rs(1)=0 Then oblog.GiveScore "",sScore,rs(0)
		'�����ȡ����˻���ɾ��
		Else
			'ֻ�����Ѿ������
			If rs(1)=1 Then oblog.GiveScore "",sScore,rs(0)
		End If
		rs.close
	Else
		tid = Split (tid ,",")
		For i = 0 To UBound(tid)
			Set rs = oblog.Execute ("select userid,istate FROM oblog_albumcomment WHERE commentid = " &tid(i))
			'����ǹ���
			If iState=1 Then
			'ֻ��������
				If rs(1)=0 Then oblog.GiveScore "",sScore,rs(0)
			'�����ȡ����˻���ɾ��
			Else
				'ֻ�����Ѿ������
				If rs(1)=1 Then oblog.GiveScore "",sScore,rs(0)
			End If
			rs.close
		Next
	End if
End Sub
Set oblog = Nothing
%>