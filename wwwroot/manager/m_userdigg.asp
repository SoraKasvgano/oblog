<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
If CheckAccess("r_user_digg")=False Then Response.Write "无权操作":Response.End
dim rs, sql
dim id,cmd,Keyword,sField,sDate1,sDate2
keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
sField=Trim(Request("Field"))
cmd=Trim(Request("cmd"))
Action=LCase(Trim(Request("Action")))
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
G_P_FileName="m_userdigg.asp?cmd=" & cmd & "&field=" & sField & "&keyword=" & keyword & "&date1=" & sDate1 & "&date2=" &sDate2
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
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">DIGG 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_userdigg.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查找：</strong></td>
      <td width="687" height="30"><select size=1 name="cmd" onChange="javascript:submit()">
          <option value=>请选择查询条件</option>
		  <option value="-1">用户推荐日志排行</option>
		  <option value="0">最新500篇用户推荐日志</option>
          <option value="1">已审核用户推荐日志</option>
          <option value="2">未通过审核的用户推荐日志</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="m_userdigg.asp">管理首页</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="m_userdigg.asp">
  <tr class="tdbg">
      <td width="120"><strong>高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">
	      <option value="author" selected>用户名称</option>
		   <option value="logid" >日志ID</option>
		  <option value="diggid" >DIGGID</option>
	      <option value="authorid" >用户ID</option>
	      <option value="ip">发表IP</option>
	      <option value="title" >标题内容</option>
	      <option value="content" >摘要内容</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="cmd" type="hidden" id="cmd" value="10">
        若为空，则查询所有</td>
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
If id<>"" Then
	If Instr(id,",") Then
		id=FilterIds(id)
	Else
		id=Int(Id)
	End If
End If
If action = "del" Or action = "best0" Or action = "best1" Or action = "pass0" Or action = "pass1" Or action = "move" Or action = "moveclass" Then
	If id = "" Then
		oblog.ShowMsg "请至少选择一个ID进行操作" , ""
	End If
End If
select Case LCase(action)
	Case "del"
		oblog.execute("DELETE FROM oblog_userdigg where diggid In ("&id & ")")
		delblogs id
		WriteSysLog "进行了用户推荐日志删除操作，目标ID："&id&"",oblog.NowUrl&"?"&OB_IIF(Request.QueryString,Request.Form)
		oblog.ShowMsg "删除成功！",""
	Case "pass0"
		oblog.execute("update oblog_userdigg Set iState=0 Where diggid In (" & id & ")")
		'进行日志更新
'		Response.Redirect "m_userdigg.asp"
		WriteSysLog "进行了用户推荐日志取消审核操作，目标ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "已设置日志为未审核状态！",""
	Case "pass1"
		oblog.execute("update oblog_userdigg Set iState=1 Where diggid In (" & id & ")")
		'进行日志更新
'		Response.Redirect "m_userdigg.asp"
		WriteSysLog "进行了用户推荐日志通过审核操作，目标ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "已设置日志为审核状态！",""
	Case "move"
		oblog.execute("update oblog_userdigg Set specialid=" & clng(Request("SpecialId")) &" Where diggid In (" & id & ")")
'		Response.Redirect "m_userdigg.asp"
		WriteSysLog "进行了用户推荐日志转移操作，目标ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "日志转移成功！",""
	Case "moveclass"
		oblog.execute("update oblog_userdigg Set classid=" & clng(Request("classid")) &" Where diggid In (" & id & ")")
'		Response.Redirect "m_userdigg.asp"
		WriteSysLog "进行了用户推荐日志分类转移操作，目标ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "日志分类转移成功！",""
	Case Else
		call main()
end select
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	Dim sQryFields, sQryTables
	sQryFields = "top 500 a.authorid,a.diggtitle,a.diggurl,a.addtime,a.diggID,a.classid,a.diggdes,a.author,a.iState,a.diggnum,a.addip,a.logid"
	sQryTables = " FROM oblog_userdigg AS a INNER JOIN oblog_log AS b ON a.logid = b.logid WHERE a.istate = 1 AND b.isdel=0 "
	select case cmd
		case -1
			sql="select " & sQryFields & sQryTables & " order by a.diggnum desc"
			sGuide=sGuide & "用户推荐日志排行"
		case 0
			sql="select " & sQryFields & sQryTables & " order by a.diggid desc"
			sGuide=sGuide & "最新500篇用户推荐日志"
		case 1
			sql="select " & sQryFields & sQryTables & " where a.iState=1 order by a.diggid desc"
			sGuide=sGuide & "通过审核的用户推荐日志"
		case 2
			sql="select " & sQryFields & sQryTables & " where a.iState=0  order by a.diggid desc"
			sGuide=sGuide & "未通过审核的用户推荐日志"
		case 10
			if Keyword="" then
				sql="select " & sQryFields & sQryTables & " order by a.diggid desc"
				sGuide=sGuide & "所有用户推荐日志"
			else
				select case sField
				case "diggid"
					if Not IsNumeric(Keyword) then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID必须是整数！</li>"
					else
						sql="select " & sQryFields & sQryTables & " where a.diggid =" & CLng(Keyword)
						sGuide=sGuide & "DIGGID等于<font color=red> " & CLng(Keyword) & " </font>的用户推荐日志"
					end If
				case "logid"
					if Not IsNumeric(Keyword) then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID必须是整数！</li>"
					else
						sql="select " & sQryFields & sQryTables & " where a.logid =" & CLng(Keyword)
						sGuide=sGuide & "LOGID等于<font color=red> " & CLng(Keyword) & " </font>的用户推荐日志"
					end if
				case "authorid"
					if Not IsNumeric(Keyword) then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID必须是整数！</li>"
					else
						sql="select " & sQryFields & sQryTables & " where a.authorid =" & CLng(Keyword)
						sGuide=sGuide & "作者ID等于<font color=red> " & CLng(Keyword) & " </font>的用户推荐日志"
					end if
				case "author"
					sql="select " & sQryFields & sQryTables & " where a.author like '%" & Keyword & "%' order by a.diggid  desc"
					sGuide=sGuide & "作者名称中含有“ <font color=red>" & Keyword & "</font> ”的用户推荐日志"
				case "ip"
					sql="select " & sQryFields & sQryTables & " where a.addip like '%" & Keyword & "%' order by a.diggid  desc"
					sGuide=sGuide & "用户推荐日志时的IP中含有“ <font color=red>" & Keyword & "</font> ”的日志"
				case "title"
					sql="select " & sQryFields & sQryTables & " where a.diggtitle like '%" & Keyword & "%' order by a.diggid  desc"
					sGuide=sGuide & "日志标题中含有“ <font color=red>" & Keyword & "</font> ”的用户推荐日志"
				case "content"
					sql="select " & sQryFields & sQryTables & " where a.diggdes like '%" & Keyword & "%' order by a.diggid  desc"
					sGuide=sGuide & "日志摘要中含有“ <font color=red>" & Keyword & "</font> ”的用户推荐日志"
				end select
			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>错误的参数！</li>"
	end select
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
'	OB_DEBUG sql,1
	If Trim(Sql)="" Then
		oblog.ShowMsg "请输入正确的查询条件！",""
	End If
	rs.Open sql,Conn,1,1
  Call oblog.MakePagebar(rs,"篇用户推荐日志")
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
  <form name="myform" method="Post" action="m_userdigg.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="0" style="word-break:break-all;">
          <%do while not rs.EOF %>
  <tr>
    <td align="center" style="background:#B3D1EA;border-bottom:1px #000 dotted;" width="30">
    	<input name='id' type='checkbox' onclick="unselectall()" id="id" value='<%=cstr(rs("diggid"))%>'>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;"><span>(<%=rs("diggnum")%>)</span><span>[<%=oblog.GetClassName(2,0,rs("classid"))%>]</span>
    	<a href="../go.asp?logid=<%=rs("logid")%>" target="_blank" style="margin:0 0 0 10px;color:#333;"><%=oblog.Filt_html(RemoveHtml(rs("diggtitle")))%></a>
    </td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="290"><a href="../go.asp?userid=<%=rs("authorid")%>" target="_blank"><font color=#0d4d89><%=rs("author")%></font></a>&nbsp;发表于
	<span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;">
	<%
		Response.write rs("addtime") & "</span>　<span style=""font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;color:#777;"">IP:" &  rs("addip")
	%></span>
	</td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="60">
		<%
			select case rs("iState")
				case 0
					Response.write "<span style=""font-weight:600;color:#f30;"">等待审核</span>"
				case 1
					Response.write "<span style=""font-weight:600;color:#090;"">通过审核</span>"
			end select
		%>
	</td>
    <td style="background:#D6EBFF;border-bottom:1px #000 dotted;" width="108">
<%
        Response.write "<a href='m_userdigg.asp?Action=Del&id=" & rs("diggid") & "' onClick='return confirm(""确定要删除此日志吗？"");'>删除</a>&nbsp;"
%>
</td>
  </tr>
  <tr>
    <td align="center" valign="top"><span style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("diggid")%></span></td>
    <td colspan="4" valign="top" style="word-wrap: break-word; word-break: break-all;color:#555;"><%=Left(oblog.Filt_html(RemoveHtml(rs("diggdes"))),200)%></td>
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
              选中本页所有日志</td>
            <td> <strong>操作：</strong>
              <input name="Action" type="radio" value="Del">
              删除&nbsp;&nbsp;
              <input name="Action" type="radio" value="pass0">
              待审&nbsp;&nbsp;
              <input name="Action" type="radio" value="pass1"">
              通过&nbsp;&nbsp;
              <input name="Action" type="radio" value="moveclass" onClick="document.myform.classid.disabled=false">
              转移&nbsp;&nbsp;
<!--               <input name="Action" type="radio" value="Move" onClick="document.myform.SpecialId.disabled=false">
              <select name="SpecialId" id="SpecialId" disabled>
              	<option value=0>取消专辑设置</option>
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
              <input type="submit" name="Submit" value="执行"> </td>
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

Sub delblogs(id)
	Dim rs,i
	Dim tid,sScore
	tid=id
	If InStr(tid,",")<0 Then
		Set rs =  Server.CreateObject("adodb.recordset")
		rs.open "SELECT b.diggnum FROM oblog_digg a INNER JOIN oblog_log b ON a.logid = b.logid WHERE a.diggID = " & Int(id),CONN,1,3
		If Not rs.Eof Then
			While Not rs.EOF
				rs(0) = 0
				rs.Update
				rs.MoveNext
			Wend
		End If
	'	oblog.Execute ("UPDATE b SET diggnum = 0  FROM oblog_digg AS a INNER JOIN oblog_log AS b ON a.logid = b.logid WHERE a.diggID =" & Int(id))
		Set rs = oblog.Execute ("SELECT COUNT(DID),authorid FROM oblog_digg WHERE diggID = " & Int(id) &" GROUP BY authorid ")
		If Not rs.Eof Then
			oblog.GiveScore "",-1*Abs(oblog.CacheScores(22))*rs(0),rs(1)
			oblog.Execute ("UPDATE oblog_user SET diggs = diggs - "&rs(0)&"  WHERE userid = " & rs(1))
		End if
		oblog.Execute ("DELETE FROM oblog_digg WHERE diggID = " & Int(id))
		rs.close
	Else
		tid = Split (tid ,",")
		For i = 0 To UBound(tid)
			Set rs =  Server.CreateObject("adodb.recordset")
			rs.open "SELECT b.diggnum FROM oblog_digg a INNER JOIN oblog_log b ON a.logid = b.logid WHERE a.diggID = " & tid(i),CONN,1,3
			If Not rs.Eof Then
				While Not rs.EOF
					rs(0) = 0
					rs.Update
					rs.MoveNext
				Wend
			End If
		'	oblog.Execute ("UPDATE b SET diggnum = 0  FROM oblog_digg AS a INNER JOIN oblog_log AS b ON a.logid = b.logid WHERE a.diggID =" & Int(id))
			Set rs = oblog.Execute ("SELECT COUNT(DID),authorid FROM oblog_digg WHERE diggID = " &  tid(i) & " GROUP BY authorid")
			If Not rs.Eof Then
				oblog.GiveScore "",-1*Abs(oblog.CacheScores(22))*rs(0),RS(1)
				oblog.Execute ("UPDATE oblog_user SET diggs = diggs - "&rs(0)&"  WHERE userid = " & rs(1))
			End if
			oblog.Execute ("DELETE FROM oblog_digg WHERE diggID = " &  tid(i))
			rs.close
		Next
	End if
End Sub
Set oblog = Nothing
%>
