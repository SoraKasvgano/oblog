<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
If CheckAccess("r_user_digg")=False Then Response.Write "无权操作":Response.End
Dim rs, sql
Dim id, cmd, Keyword, sField
Keyword = Trim(Request("keyword"))
If Keyword <> "" Then Keyword = oblog.filt_badstr(Keyword)
sField = Trim(Request("Field"))
cmd = Trim(Request("cmd"))
Action = Trim(Request("Action"))
If cmd = "" Then
    cmd = 0
Else
    cmd = CLng(cmd)
End If
G_P_FileName = "m_digg.asp?cmd=" & cmd & "&Field=" & sField & "&keyword=" & Keyword
%>
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
		<li class="main_top_left left"><%If cmd = 0 Then %>DIGG 记 录 管 理<%else%>反 映 问 题 管 理<%End if%></li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_digg.asp?cmd=<%=cmd%>" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查找：</strong></td>
      <td width="687" height="30">
        <select name="Field" id="Field">
		<%If cmd = 1 Or cmd = 3 Then%>
            <option value="author">反映人名称</option>
            <option value="ip">反映人ip</option>
            <option value="userid">目标用户ID</option>
			<option value="logid">日志ID</option>
		<input type="hidden" name="cmd" value="3">
		<%Else%>
            <option value="author">推荐人名称</option>
            <option value="ip">推荐人ip</option>
            <option value="userid">目标用户ID</option>
			<option value="diggid">DIGGID</option>
			<option value="logid">日志ID</option>
			 <input type="hidden" name="cmd" value="2">
		<%End if%>
        </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit"  value=" 搜索 ">&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  </form>
  <form action="m_digg.asp" name="form2" method="get">
  <tr class="tdbg">
      <td width="100"><strong>数据清理：</strong></td>
    <td>
            按IP清理&nbsp;
            <input name="ip" type="text" size="20" maxlength="30">
            <input type="checkbox"  name="chkIp" value="1" checked>是否将该IP加入到黑名单
            <input type="hidden" name="action" value="clearip">
          <input type="submit"  value="清理" />
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
	Case "Del"
		Call Optcomment("1")
	Case "clearip"
		Call Optcomment("2")
	Case "pass0"
		id = GetLogID
		oblog.execute("update oblog_log Set passcheck=0 Where logid In (" & id & ")")
		'进行日志更新
		DoUpdatelog id
		WriteSysLog "进行了日志取消审核操作，目标日志ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "已设置日志为未审核状态！",""
	Case "pass1"
		id = GetLogID
		oblog.execute("update oblog_log Set passcheck=1 Where logid In (" & id & ")")
		'进行日志更新
		DoUpdatelog id
		WriteSysLog "进行了日志通过审核操作，目标日志ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "已设置日志为审核状态！",""
	Case "Dellog"
		
		id = GetLogID
		Call DelScore(id)
		oblog.execute("update oblog_log Set isdel=1 where logid In ("&id & ")")
		'删除日志文件!
		delblogs id
		Call Optcomment("1")
		WriteSysLog "进行了日志删除操作（放入回收站），目标日志ID："&id&"",oblog.NowUrl&"?"&OB_IIF(Request.QueryString,Request.Form)
		oblog.ShowMsg "删除成功！",""
	Case Else
	    Call main
End select
If ErrMsg<>"" Then
    Call WriteErrMsg
End If

Sub Main()
	Dim SQL3
	sql = "select top 500 a.userid,b.diggtitle,a.addtime,did,username,a.addip,diggtype,a.diggid,isguest,b.authorid,b.logid From oblog_digg a LEFT JOIN oblog_userdigg b ON a.diggid = b.diggid"
    select Case cmd
        Case 0,""
        	sql= Sql & " WHERE a.diggtype = -1 Order By did desc"
            sGuide = sGuide & "所有推荐记录"
		Case 1
			sql = "select top 500 a.userid,b.topic,a.addtime,did,username,a.addip,diggtype,a.diggid,isguest,a.authorid,A.logid From oblog_digg a LEFT JOIN oblog_log b ON a.logid = b.logid WHERE a.diggtype > -1"
        	sql= Sql & " Order By did desc"
            sGuide = sGuide & "所有反映问题"
        Case 2,3
            If Keyword = "" Then
            	ErrMsg="错误：关键字不能为空！"
                Exit Sub
            Else
				If cmd = 3 Then
					sql = "select top 500 a.userid,b.topic,a.addtime,did,username,a.addip,diggtype,a.diggid,isguest,a.authorid,a.logid From oblog_digg a LEFT JOIN oblog_log b ON a.logid = b.logid"
					SQL3 = " AND a.diggtype > -1"
				Else
					SQL3 = " AND a.diggtype = -1"
				End if
                select Case sField
	                Case "author"
	                    sql= Sql & " Where username like '%" & Keyword&"%' "&SQL3&" order by did desc"
	                    sGuide = sGuide & "推荐者名称中还有含有<font color=red> " & Keyword & " </font>的记录"
	                Case "userid"
	                    sql= Sql & " Where a.authorid =" & Int(Keyword)&" "&SQL3&" order by did desc"
	                    sGuide = sGuide & "作者ID为<font color=red> " & Keyword & " </font>接受到的记录"
	                Case "ip"
	                    Sql= Sql & " Where a.addip='" & Keyword&"' "&SQL3&" order by did desc"
	                    sGuide = sGuide & "作者ip为<font color=red> " & Keyword & " </font>的记录"
	                Case "diggid"
	                    sql= Sql & " Where a.diggid =" & Int(Keyword)&" order by did desc"
	                    sGuide = sGuide & "DIGGID为<font color=red> " & Keyword & " </font>的记录"
	                Case "logid"
	                    sql= Sql & " Where a.logid =" & Int(Keyword)&" "&SQL3&" order by did desc"
	                    sGuide = sGuide & "日志ID为<font color=red> " & Keyword & " </font>的记录"
                End select
            End If
        Case Else
        	Exit sub
    End Select
'	OB_DEBUG Sql,1
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, conn, 1, 1
    Call oblog.MakePageBar(rs, "条记录")
    rs.Close
    Set rs = Nothing
End Sub
Sub showContent()
'	On Error Resume Next
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
  <form name="myform" id="myform" method="post" action="m_digg.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
<style type="text/css">
<!--
td {padding:3px 0!important;}
-->
</style>
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="0" style="word-wrap: break-word; word-break: break-all;">
          <%do while not rs.EOF %>
  <tr>
    <td align="center" style="background:#B3D1EA;border-bottom:1px #000 dotted;" width="30"><input type="checkbox" name="chkOne" id="id" value='<%=rs("did")%>'>
	<%If rs("diggtype") = -1 Then %>
	<input type="hidden" name="authorid" id="authorid" value='<%=rs("authorid")%>'><%End if%></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;"><a href="../go.asp?logid=<%=OB_IIF(rs("logid"),0)%>" target="_blank" style="margin:0 0 0 10px;color:#333;"><%
	Dim Topic
	Topic = oblog.filt_html(RemoveHtml(Left(rs(1),20)))
	If IsNull(rs(1)) Then
		Response.Write "日志已被删除"
	Else
		Response.Write Topic
	End if
	%></a></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;color:#666;" width="300"><font color=#0d4d89><a href="../go.asp?userid=<%=rs("userid")%>" target="_blank"><%=OB_IIF(rs("username"),"未记录")%></a></font><%If rs("isguest") = 1 Then Response.Write "(游客)"%>&nbsp;发表于&nbsp;<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;"><%=rs("addtime")%></span>	　<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;">IP:<%=rs("addip")%></span></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="110">目标用户ID:<a href="../go.asp?userid=<%=OB_IIF(rs("authorid"),0)%>" target="_blank" style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=OB_IIF(rs("authorid"),"无")%></a></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="40">&nbsp;
</td>
  </tr>
  <tr>
    <td align="center" valign="top"><span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("did")%></span></td>
    <td colspan="4" valign="top" style="word-wrap: break-word; word-break: break-all;font-weight:600;color:#f00;"><%
	If rs("diggtype") >-1 Then Response.Write oblog.CacheReport(rs("diggtype"))
	%></td>
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
                <input type="checkbox" name="chkAll" id="chkAll" onClick="javascript:CheckAll(this.form);">全选
                &nbsp;&nbsp;&nbsp;&nbsp;
                操作方式:
				<input type="radio" name="Action" value="Del">删除
              <input name="Action" type="radio" value="Dellog">
              删除日志&nbsp;&nbsp;
              <input name="Action" type="radio" value="pass0">
              待审日志&nbsp;&nbsp;
              <input name="Action" type="radio" value="pass1"">
              审核日志&nbsp;&nbsp;
		<input type="hidden" id="cmd" name ="cmd" value="<%=cmd%>" />
                <input type="submit" value="开始操作" name="submit">
				<br />
				&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>（对日志的操作是指此条记录所关联的某篇日志）</font>
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
    '删除/待审/审核
    Dim sOpt,sIds,sChkIp,sIP,aIds,l_aIds,i,RSDIGG,l_lids,L_id
    sIP=Request("ip")
	sChkIp=Request("chkip")
    sIds = FilterIds(Request("chkOne"))
    aIds = FilterIds(Request("authorid"))
    sOpt = Request("opt")
    '按ID批量删除
    If sMod="1" Then
		If Request.QueryString <> "" Then Exit Sub
	    If sIds = "" Then oblog.ShowMsg "未选择操作问题id", ""
		WriteSysLog "进行了DIGG记录（用户反映问题）删除操作，目标ID："&sIds&"",""
		sql = "Delete From oblog_digg Where did In (" & sIds & ")"
		
	   
		l_lids= Split(sIds,",")
		For i=0 To UBound(l_lids)
		L_id=oblog.execute("select top 1 logid from oblog_digg where did="&l_lids(i))(0)
		oblog.execute("update oblog_log set diggNum = diggNum -1 where logid="&L_id)
		oblog.execute("update oblog_userdigg set diggNum = diggNum -1 where logid="&L_id)
		Next 
		 oblog.Execute sql
		l_aIds = Split (aIds,",")
		For i = 0 To UBound(l_aIds)
			oblog.GiveScore "",-1*Abs(oblog.CacheScores(22)),l_aIds(i)
			Oblog.Execute ("UPDATE oblog_user SET diggs = diggs - 1 WHERE userid = " & l_aIds(i))
		Next
	Else
		sIp=oblog.filt_badstr(sIp)
		'OB_Debug Request("ip"),1
		If sIp="" Then Exit Sub
		If Not IsObject(CONN) Then link_database
		Set RSDIGG = Server.CreateObject("ADODB.RecordSet")
		RSDIGG.open "SELECT authorid,diggtype From oblog_digg Where addIp='" & sIp & "'",CONN,1.3
		If Not RSDIGG.EOF Then
			While Not RSDIGG.EOF
				If RSDIGG(1) = -1 Then
					oblog.GiveScore "",-1*Abs(oblog.CacheScores(22)),RSDIGG(0)
					Oblog.Execute ("UPDATE oblog_user SET diggs = diggs - 1 WHERE userid = " & RSDIGG(0))
				End if
				RSDIGG.DELETE
				RSDIGG.MoveNext
			Wend
			If sChkIp = "1" And oblog.ChkWhiteIP(sIP) = False Then
				'加入黑名单
				oblog.KillIP(sIP)
			End If
		End If
		WriteSysLog "进行了DIGG记录（用户反映问题）清理操作，目标IP："&sIp&"",oblog.NowUrl&"?"&Request.QueryString
	End If
    oblog.ShowMsg "操作成功!", ""
End Sub

'更新日志
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
			'清理文件记录
			'Call oblog.DeleteFiles(logid)
			'真实域名需要重新整理文件数据
			'物理文件即时删除
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
			'更新计数器,删除积分
			If rst("isdel")=1 Then
				Call Tags_UserDelete(logid(i))
				Call OBLOG.log_count(uid,logid(i),sid,cid,"-")
			End If
			'--------------------------------------------
			blog.userid = uid
			blog.Update_Subject uid
			blog.Update_index 0
			blog.Update_newblog (uid)
		End If
		rst.Close
  	Next
	Set blog = Nothing
	Set fso = Nothing
	Set rst = Nothing
End Sub

Function GetLogID()
	Dim RS,tmpid,SID
	SID = FilterIds(Request("chkOne"))
	If sid="" Or isnull(sid) Then oblog.ShowMsg "未选择操作问题id", ""
	Set RS = oblog.Execute ("SELECT logid FROM oblog_digg WHERE did IN ("&SID&")")
	If Not RS.Eof Then
		While Not RS.Eof
			tmpid = tmpid  & ","&RS(0)
			RS.MoveNext
		Wend
		tmpid = FilterIds(tmpid)
	End If
	GetLogID = tmpid
End Function

Sub DelScore(id)
	Dim rs,i
	Dim tid,sScore
	tid=id
	'删除日志时，将删除该日志所获得的所有积分,并且进行积分惩罚
	If InStr(tid,",")<0 Then
		Set rs = oblog.Execute ("select userid,scores FROM oblog_log WHERE logid = " &tid)
		If Not rs.Eof Then
			sScore=-1*(rs(1)+CLng(oblog.CacheScores(4)))
			If IsNull(sScore) Then sScore = -1*(CLng(oblog.CacheScores(4)))
			oblog.GiveScore "",sScore,rs(0)
		End if
		rs.close
	Else
		tid = Split (tid ,",")
		For i = 0 To UBound(tid)
			Set rs= oblog.execute ("select userid,scores FROM oblog_log WHERE logid = " &tid(i))
			If Not rs.Eof Then
				sScore=-1*(rs(1)+CLng(oblog.CacheScores(4)))
				If IsNull(sScore) Then sScore = -1*(CLng(oblog.CacheScores(4)))
				oblog.GiveScore "",sScore,rs(0)
			End if
			rs.close
		Next
	End if
End Sub
Set oblog = Nothing
%>