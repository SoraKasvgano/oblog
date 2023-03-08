<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_album_comment")=False Then Response.Write "无权操作":Response.End
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
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">相 册 评 论 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_album_comments.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查找：</strong></td>
      <td width="687" height="30">
        <select name="Field" id="Field">
            <option value="author">评论人名称</option>
            <option value="ip">评论人ip</option>
            <option value="userid">用户ID</option>
            <option value="topic">评论标题</option>
            <option value="content">评论内容</option>
        </select>
      <input type="hidden" name="cmd" value="2">
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit"  value=" 搜索 ">&nbsp;&nbsp;&nbsp;&nbsp;<a href="m_album_comments.asp">最新评论</a>|&nbsp;&nbsp;&nbsp;&nbsp;<a href="m_album_comments.asp?cmd=1">待审核评论</a></td>
    </tr>
  </form>
  <form action="m_album_comments.asp" name="form2" method="get">
  <tr class="tdbg">
      <td width="100"><strong>数据清理：</strong></td>
    <td>
            按IP清理评论&nbsp;
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
            sGuide = sGuide & "所有评论"
        Case 1
        	sql= Sql & " Where iState=0 Order By commentid desc"
        	sGuide = sGuide & "待审核评论"
        Case 2
            If Keyword = "" Then
            	ErrMsg="错误：关键字不能为空！"
                Exit Sub
            Else
                select Case sField
	                Case "author"
	                    sql= Sql & " Where comment_user like '%" & Keyword&"%' order by commentid desc"
	                    sGuide = sGuide & "评论者名称中还有含有<font color=red> " & Keyword & " </font>的评论"
	                Case "userid"
	                    sql= Sql & " Where userid =" & Int(Keyword)&" order by commentid desc"
	                    sGuide = sGuide & "被评论者ID为<font color=red> " & Keyword & " </font>接受到的评论"
	                Case "topic"
	                    sql= Sql & " Where commenttopic like '%" & Keyword & "%' order by commentid desc"
	                    sGuide = sGuide & "标题中含有“ <font color=red>" & Keyword & "</font> ”的评论"
	                Case "ip"
	                    Sql= Sql & " Where addip='" & Keyword&"' order by commentid desc"
	                    sGuide = sGuide & "作者ip为<font color=red> " & Keyword & " </font>的评论"
	                Case "content"
	                    sql= Sql & " Where comment like '%" & Keyword&"%' order by commentid desc"
	                    sGuide = sGuide & "评论内容中包含<font color=red> " & Keyword & " </font>的评论"
                End select
            End If
        Case Else
        	Exit sub
    End select
    'Response.Write Sql
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, conn, 1, 1
    Call oblog.MakePageBar(rs, "篇评论")
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
  <form name="myform" id="myform" method="post" action="m_album_comments.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
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
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;color:#666;" width="300"><font color=#0d4d89><%=rs("comment_user")%></font><%If rs("isguest") = 1 Then Response.Write "(游客)"%>&nbsp;发表于&nbsp;<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;"><%=rs("addtime")%></span>	　<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;">IP:<%=rs("addip")%></span></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="110">目标用户ID:<a href="../go.asp?userid=<%=rs("userid")%>" target="_blank" style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("userid")%></a></td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="40">&nbsp;
<%If rs("iState")="1" Then %><span style="font-weight:600;color:#090;">已审</span><%Else%><span style="font-weight:600;color:#f30;">待审</span><%End If%>
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
                <input type="checkbox" name="chkAll" id="chkAll" onClick="javascript:CheckAll(this.form);">全选
                &nbsp;&nbsp;&nbsp;&nbsp;
                操作方式:
				<input type="radio" name="opt" value="3">删除
				<input type="radio" name="opt" value="2">取消审核
                <input type="radio" name="opt" value="1">通过审核&nbsp;&nbsp;
				<input type="hidden" value="batchopt" name="action">
                <input type="submit" value="开始操作" name="submit">
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
    '删除/待审/审核
    Dim sOpt,sIds,sChkIp,sIP
    sIP=Request("ip")
	sChkIp=Request("chkip")
    sIds = FilterIds(Request("chkOne"))
    sOpt = Request("opt")
    '按ID批量删除
    If sMod="1" Then
	    If sOpt = "" Or sIds = "" Then Call main(): Exit Sub
	    '取单一userid
	    'Set rstUser = oblog.Execute("select b.userid,count(a.commentid) From oblog_albumcomment a,(select userid From oblog_albumcomment Where commentId In (" & sIds & ") Group By userid) b Where a.userid=b.userid")
	    Set rstUser = oblog.Execute("select userid From oblog_albumcomment Where commentId In (" & sIds & ") Group By userid")
	    '取单一主logid
	    Set rstBlog = oblog.Execute("select mainid From oblog_albumcomment Where commentid In (" & sIds & ") Group By mainid")
		PassScore sIds,sOpt
	    select Case sOpt
	        Case "1"
	            '先给分，后处理，以获得初始状态
				WriteSysLog "进行了相册评论通过审核操作，目标评论ID："&sIds&"",""
	            sql = "Update oblog_albumcomment Set iState=1 Where commentId In (" & sIds & ")"
	        Case "2"
				WriteSysLog "进行了相册评论取消审核操作，目标评论ID："&sIds&"",""
	            sql = "Update oblog_albumcomment Set iState=0 Where commentId In (" & sIds & ")"
	        Case "3"
				WriteSysLog "进行了相册评论删除操作，目标评论ID："&sIds&"",""
	            sql = "Delete From oblog_albumcomment Where commentId In (" & sIds & ")"
	        Case Else
	            rstUser.Close
	            Set rstUser = Nothing
	            Exit Sub
	    End select
	    oblog.Execute sql
		MSG = "评论管理操作成功!"
	Else
		MSG = "清理评论成功"
		sIp=oblog.filt_badstr(sIp)
		'OB_Debug Request("ip"),1
		If sIp="" Then Exit Sub
		'Set rstUser = oblog.Execute("select userid,count(commentid)  From oblog_albumcomment Where addip='" & sIP & "' Group By userid")
		Set rstUser = oblog.Execute("select userid From oblog_albumcomment Where addip='" & sIP & "' Group By userid")
		Set rstBlog = oblog.Execute("select mainid From oblog_albumcomment Where addip='" & sIP & "' Group By mainid")
		oblog.Execute ("Delete From oblog_albumcomment Where addIp='" & sIp & "'")
		If sChkIp = "1" And oblog.ChkWhiteIP(sIP) = False Then
			'加入黑名单
			oblog.KillIP(sIP)
	    End If
		WriteSysLog "进行了相册评论清理操作，目标评论IP："&sIp&"",oblog.NowUrl&"?"&Request.QueryString
	End If
    Dim blogcomments,allComments
    '对日志评论重新进行计数
    Do While Not rstUser.EOF
        '更新用户积分
        Set rstComment=oblog.Execute("select Count(commentid) From oblog_comment Where istate=1 AND userid=" & rstUser(0))
		allComments =  rstComment(0)
		rstComment.Close
        Set rstComment=oblog.Execute("select Count(commentid) From oblog_albumcomment Where istate=1 AND userid=" & rstUser(0))
		allComments = allComments + rstComment(0)
        '评论数目
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
        '重新计算评论数目
        oblog.Execute ("update [oblog_album] set commentnum=" & blogcomments  & " Where fileid=" & rstBlog(0))
        rstBlog.MoveNext
    Loop
    rstBlog.Close
    Set rstUser = Nothing
    Set rstBlog = Nothing
    oblog.ShowMsg MSG, ""
End Sub
'iState=1 通过审核;2取消审核;3删除
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
		'如果是过审
		If iState=1 Then
			'只处理待审的
			If rs(1)=0 Then oblog.GiveScore "",sScore,rs(0)
		'如果是取消审核或者删除
		Else
			'只处理已经过审的
			If rs(1)=1 Then oblog.GiveScore "",sScore,rs(0)
		End If
		rs.close
	Else
		tid = Split (tid ,",")
		For i = 0 To UBound(tid)
			Set rs = oblog.Execute ("select userid,istate FROM oblog_albumcomment WHERE commentid = " &tid(i))
			'如果是过审
			If iState=1 Then
			'只处理待审的
				If rs(1)=0 Then oblog.GiveScore "",sScore,rs(0)
			'如果是取消审核或者删除
			Else
				'只处理已经过审的
				If rs(1)=1 Then oblog.GiveScore "",sScore,rs(0)
			End If
			rs.close
		Next
	End if
End Sub
Set oblog = Nothing
%>