<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
If CheckAccess("r_group_user")=False Then Response.Write "无权操作":Response.End
dim rs, sql
dim UserID,cmd,Keyword,sField
dim str
keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
sField=Trim(Request("Field"))
cmd=Trim(Request("cmd"))
Action=Trim(Request("Action"))
UserID=Trim(Request("UserID"))


if cmd="" then
	cmd=0
else
	cmd=CLng(cmd)
end if

G_P_FileName="m_team.asp?cmd=" & cmd
if sField<>"" then
	G_P_FileName=G_P_FileName&"&Field="&sField
end if
if keyword<>"" then
	G_P_FileName=G_P_FileName&"&keyword="&keyword
	cmd=10
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--<%=oblog.CacheConfig(69)%>管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left"><%=oblog.CacheConfig(69)%>管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" id="table1">
  <form name="form1" action="m_team.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查找<%=oblog.CacheConfig(69)%>：</strong></td>
      <td width="687" height="30">
		<select size=1 name="cmd" onChange="javascript:submit()">
          <option value=>请选择查询条件</option>
		  <option value="0">等待认证的<%=oblog.CacheConfig(69)%></option>
          <option value="1">被锁住的<%=oblog.CacheConfig(69)%></option>
          <option value="2">最后注册的50个<%=oblog.CacheConfig(69)%></option>
          <option value="3">积分最高TOP10</option>
          <option value="4">积分最低的10个<%=oblog.CacheConfig(69)%></option>
          <option value="5">推荐<%=oblog.CacheConfig(69)%></option>

        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="m_team.asp?cmd=2"><%=oblog.CacheConfig(69)%>管理首页</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="m_team.asp">
  <tr class="tdbg">
    <td width="120"><strong><%=oblog.CacheConfig(69)%>高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">
      <option value="teamid" selected><%=oblog.CacheConfig(69)%>ID</option>
	  <option value="teamname" selected><%=oblog.CacheConfig(69)%>名</option>
      <option value="UserID" ><%=oblog.CacheConfig(70)%>ID</option>
	  <option value="username" ><%=oblog.CacheConfig(70)%>名</option>

      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="cmd" type="hidden" id="cmd" value="10">
	  若为空，则查询所有<%=oblog.CacheConfig(69)%></td>
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
Dim s1,actionStr,teamid
s1=Request("s1")
teamid=Request("teamid")
If teamid<>"" Then
	If Instr(teamid,",") Then
		teamid=FilterIds(teamid)
	Else
		teamid=Int(teamid)
	End If
End If
'OB_DEBUG teamid,1
select Case LCase(Action)
	Case "modifystate","pass0","pass1"
		If Action = "pass0" Then s1 = "2"
		If Action = "pass1" Then s1 = "1"
			select Case s1
				Case "1"
					actionStr="通过审核"
					Call TeamScore(teamid,1)
					oblog.Execute("Update oblog_team Set istate=3 Where teamid IN (" & teamid &") ")
					WriteSysLog "进行了通过审核"&oblog.CacheConfig(69)&"操作，目标"&oblog.CacheConfig(69)&"ID："&teamid&"",oblog.NowUrl&"?"&Request.QueryString
				Case "2"
					actionStr="锁定"
					oblog.Execute("Update oblog_team Set istate=2 Where teamid IN (" & teamid &") ")
					WriteSysLog "进行了锁定"&oblog.CacheConfig(69)&"操作，目标"&oblog.CacheConfig(69)&"ID："&teamid&"",oblog.NowUrl&"?"&Request.QueryString
				Case "3"
					actionStr="解除锁定"
					oblog.Execute("Update oblog_team Set istate=3 Where teamid IN (" & teamid &") ")
					WriteSysLog "进行了解除锁定"&oblog.CacheConfig(69)&"操作，目标"&oblog.CacheConfig(69)&"ID："&teamid&"",oblog.NowUrl&"?"&Request.QueryString
			End select
			oblog.ShowMsg "群组" & actionStr & "成功",""
	Case "best"
		call best()
	Case "delteam"
		oblog.Execute("DELETE FROM oblog_team Where teamid in (" & teamid &" )")
		oblog.Execute("DELETE FROM oblog_teamusers Where teamid in (" & teamid &" )")
		WriteSysLog "进行了删除"&oblog.CacheConfig(69)&"操作，目标"&oblog.CacheConfig(69)&"ID："&teamid&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "群组删除成功",""
	Case "modify"
		Call modify()
	case "modifyteaminfo"
		call modifyteaminfo()
	Case else
		call main()
end select
if FoundErr=true then
	call WriteErrMsg()
end If

Sub TeamScore(teamid,istate)
	Dim rs
	Set rs = oblog.Execute ("select createrid,istate FROM oblog_team WHERE teamid IN (" & teamid &") ")
	If rs(1) = 1 Then
		If istate = 3 Then
			oblog.GiveScore "" ,oblog.CacheScores(12),rs(0)
		End If
	ElseIf rs(1) = 3 Then
		If istate = 1 Then
			oblog.GiveScore "" ,oblog.CacheScores(12),rs(0)
		End If
	End if
	rs.close
	Set rs=Nothing
End Sub

sub main()

	sGuide=""
	select case cmd
		case 0
			sql="select top 500 * from oblog_team Where istate=1 order by teamid desc"
			sGuide=sGuide & "等待管理认证的" & oblog.CacheConfig(69)
		case 1
			sql="select top 500 * from oblog_team Where istate=2 order by teamid desc"
			sGuide=sGuide & "所有被锁住的" & oblog.CacheConfig(69)
		case 2
			sql="select top 500 * from oblog_team Where istate>0 order by teamid desc"
			sGuide=sGuide & "最后注册的500个" & oblog.CacheConfig(69)
		case 3
			sql="select top 500 * from oblog_team Where istate=3 order by teamscore desc"
			sGuide=sGuide & "积分最高的前500个" & oblog.CacheConfig(69)
		case 4
			sql="select top 500 * from oblog_team Where istate=3 order by teamscore"
			sGuide=sGuide & "积分最少的10个" & oblog.CacheConfig(69)
		case 5
			sql="select top 500 * from oblog_team Where isbest=1 order by teamscore"
			sGuide=sGuide & "推荐" & oblog.CacheConfig(69)
		case 10
			if Keyword="" then
				sql="select top 500 * from oblog_team order by teamid Desc"
				sGuide=sGuide & "所有" & oblog.CacheConfig(69)
			else
				select case LCase(sField)
				case "userid"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>"&oblog.CacheConfig(70)&"ID必须是整数！</li>"
					else
						sql="select * from oblog_team where managerid=" & CLng(Keyword)
						sGuide=sGuide & oblog.CacheConfig(70) & "ID等于<font color=red> " & CLng(Keyword) & " </font>的" & oblog.CacheConfig(69)
					end if
				case "username"
					sql="select * from oblog_team where managername like '%" & Keyword & "%'"
					sGuide=sGuide & oblog.CacheConfig(70) &"用户名中含有“ <font color=red>" & Keyword & "</font> ”的" & oblog.CacheConfig(69)
				case "teamname"
					sql="select * from oblog_team where t_name like '%" & Keyword & "%'"
					sGuide=sGuide & oblog.CacheConfig(69) &"名称中含有“ <font color=red>" & Keyword & "</font> ”的" & oblog.CacheConfig(69)
				case "teamid"
					sql="select * from oblog_team where teamid=" & CLng(Keyword)
					sGuide=sGuide & oblog.CacheConfig(69) &"ID等于“ <font color=red>" & Keyword & "</font> ”的" & oblog.CacheConfig(69)
				end select
			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>错误的参数！</li>"
	end select
	If sGuide="" Then sGuide="群组管理"
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'Response.write sql
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
	Call oblog.MakePageBar(rs,"个" & oblog.CacheConfig(69))
	%>
	 </table>
		</div>
	</div>
	<%
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i
    i=0
%>
<style type="text/css">
<!--
.border tr td {padding:3px 0!important;}
-->
</style>
  <form name="myform" method="Post" action="" onsubmit="return confirm('确定要执行选定的操作吗？');">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border" id="table3">
          <tr class="title">
			<td width="30" align="center"><strong>选中</strong></td>
            <td width="44" align="center"><strong>ID</strong></td>
            <td width="90" align="center"><strong><%=oblog.CacheConfig(69)%>LOGO</strong></td>
            <td width="140" align="center"><strong><%=oblog.CacheConfig(69)%>名 <%=oblog.CacheConfig(70)%> 申请时间</strong></td>
            <td width="60" align="center"><strong>会员数</strong></td>
            <td width="50" align="center"><strong>主贴</strong></td>
            <td width="50" align="center"><strong>回复</strong></td>
            <td align="center"><strong>申请说明</strong></td>
            <td  width="70" align="center" ><strong>状态</strong></td>
			<td  width="80" align="center" ><strong>操作</strong></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg">
		  <td align="center">
		  <input type="checkbox" id="teamid" name = "teamid" value="<%=rs("teamid")%>"/></td>
            <td style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;" align="center"><%=rs("teamid")%></td>
            <td align="center"><a href="<%=ProIco(rs("t_ico"),3)%>" target="_blank" title="点击查看该图"><img src="../<%=rs("t_ico")%>" align="absmiddle" style="width:80px;height:60px;border:0;"></a></td>
            <td><span style="display:block;padding:0 0 0 8px!important;"><a href="../group.asp?gid=<%=rs("teamid")%>" target="_blank"><%=rs("t_name")%></a></span><span style="display:block;padding:0 0 0 8px!important;color:#217dbd;">组长：<a href="../go.asp?userid=<%=rs("managerid")%>" target="_blank"><%=rs("managername")%></a></span><span style="display:block;color:#999;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:10px;padding:0 0 0 8px!important;"><%=rs("createtime")%></span></td>
            <td align="center" style="font-weight:600;color:#217dbd;">
			<%=rs("icount0")%>人
            </td>
        	<td align="center" style="font-weight:600;color:#f00;">
            <%=rs("icount1")%>
			</td>
        	<td align="center" style="font-weight:600;color:#090;">
            <%=rs("icount2")%>
			</td>
            <td valign="top">
			<span style="padding:6px;"><%=oblog.Filt_html(OB_IIF(rs("intro"),""))%></span>
			</td>
            <td  align="center">
            <%select case cint(rs("istate"))
            	case 1
            		str="<span style=""color:#f60;font-weight:600;"">待审</span>"
            	case 2
            		str="<span style=""color:#f00;font-weight:600;"">锁定</span>"
            	case 3
            		str="<span style=""color:#090;font-weight:600;"">审核</span>"
            end select
            Response.write str
            %>

            <%select case cint(rs("istate"))
            	case 1%>
            	<a href="m_team.asp?action=modifystate&s1=1&teamid=<%=rs("teamid")%>" onClick="return confirm('确定要批准通过此<%=oblog.CacheConfig(69)%>吗？');">通过</a>
            	<%case 2%>
            	<a href="m_team.asp?action=modifystate&s1=3&teamid=<%=rs("teamid")%>" onClick="return confirm('确定要解锁此<%=oblog.CacheConfig(69)%>吗？');">解锁</a>
            	<%case 3%>
            	<a href="m_team.asp?action=modifystate&s1=2&teamid=<%=rs("teamid")%>" onClick="return confirm('确定要锁定此<%=oblog.CacheConfig(69)%>吗？');">锁定</a>
            <%end select%>
            </td>
			<td align="center"><a href="?action=modify&teamid=<%=rs("teamid")%>">修改</a>&nbsp;&nbsp;<a href="m_team.asp?action=delteam&teamid=<%=rs("teamid")%>" onClick="return confirm('确定要删除此<%=oblog.CacheConfig(69)%>吗？');">删除</a></td>
          </tr>
          <%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext

Loop
Response.write "</table>"%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="140" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页所有内容</td>
            <td> <strong>操作：</strong>
              <input name="action" type="radio" value="delteam">
              删除&nbsp;&nbsp;
              &nbsp;&nbsp;
              <input name="action" type="radio" value="pass0">
              锁定&nbsp;&nbsp;
              &nbsp;&nbsp;
              <input name="action" type="radio" value="pass1">
              审核&nbsp;&nbsp;
              &nbsp;&nbsp;
              <input type="submit" name="Submit" value="执行"> </td>
  </tr>
</table>
</form>
<%
end Sub
Sub modify()
	set rs=oblog.execute("select * from oblog_team where teamid="&teamid&"")
	ReCountTeamInfo(teamid)
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">修改<%=oblog.CacheConfig(69)%>信息</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">

<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" id="table1">
  <form name="oblogform" method="post" action="?action=modifyteaminfo&teamid=<%=teamid%>">
    <tr class="tdbg">
      <td width="100"><%=oblog.CacheConfig(69)%>创建者：</td>
      <td><input type="text" name="creatername" size="30" value="<%=rs("creatername")%>" /></td>
    </tr>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>管理员：</td>
      <td><input type="text" name="managername" size="30" value="<%=rs("managername")%>" /></td>
    </tr>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>名称：</td>
      <td><input type="text" name="t_name" size="30" value="<%=rs("t_name")%>" /></td>
    </tr>
	  <%If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" Then%>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>域名：</td>
      <td><input name="team_domain" type="text" value="<%=rs("t_domain")%>" size=10 maxlength=20 /> <select name="team_domainroot" ><%=oblog.type_domainroot(rs("t_domainroot"),1)%></select></td>
    </tr>
	  <%End if%>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>标记图片<br />(120*90)</td>
      <td><div class="user_face"><img id="imgIcon" width=120 height=90 src="<%=oblog.filt_html(ProIco(rs("t_ico"),3))%>">
		<p><iframe id="d_file" frameborder="0" src="../upload.asp?tMode=8&re=&teamid=<%=teamId%>" width="300" height="80" scrolling="no"></iframe>
			<br/>只支持jpg,gif,png,小于200k,默认尺寸为120*90<br/>
			图片地址：<input name="ico"  type="text" value="<%=oblog.filt_html(rs("t_ico"))%>" size="70" maxlength="200" / >
			<br/>你可以直接输入一个有效的图片地址,也可以在这里直接选择一个系统可用的图片</p></div></td>
    </tr>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>类别：</td>
      <td><select name="classid" id="classid" ><%=oblog.show_class("log",rs("classid"),2)%></select></td>
    </tr>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>标签：</td>
      <td><input type="text" name="tags" size="50" value="<%=rs("t_tags")%>">(最多支持5个，以逗号间隔)</td>
    </tr>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>加入条件：</td>
      <td>
			<input type="radio" name="t1" value="-1" <%If rs("joinlimit")=-1 Then Response.Write " checked" End If%>>任意加入
			<input type="radio" name="t1" value="0" <%If rs("joinlimit")=0 Then Response.Write " checked" End If%>>申请加入
			<input type="radio" name="t1" value="1" <%If rs("joinlimit")=1 Then Response.Write " checked" End If%>>仅可邀请<br/>
			<input type="radio" name="t1" value="2"  <%If rs("joinlimit")=2 Then Response.Write " checked" End If%>>积分限制，需大于<input type=text name="t2" size=5 maxlength=8 value="<%=rs("joinscores")%>">积分才能申请
		</td>
    </tr>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>访问权限：</td>
      <td>
			<input type="radio" name="t4" value="-1" <%If rs("viewlimit")=-1 Then Response.Write " checked" End If%>>无限制
			<input type="radio" name="t4" value="0" <%If rs("viewlimit")=0 Then Response.Write " checked" End If%>><%=oblog.CacheConfig(69)%>成员可见 <br/>
			<input type="radio" name="t4" value="1" <%If rs("viewlimit")=1 Then Response.Write " checked" End If%>>凭密码访问，密码<input type=text name="t5" size=20 maxlength=20 value="">（不修改请留空）
		</td>
    </tr>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>申请说明/简介：</td>
      <td><textarea rows="5" name="intro" cols="45"><%=rs("intro")%></textarea></td>
    </tr>
    <tr class="tdbg">
      <td><%=oblog.CacheConfig(69)%>状态：</td>
      <td>
			<input type="hidden" name="istate_0" id="istate_0" value = "<%=rs("istate")%>" />
			<input type="radio" name="istate" value="1" <%If rs("istate")=1 Then Response.Write " checked" End If%>>待审
			<input type="radio" name="istate" value="2" <%If rs("istate")=2 Then Response.Write " checked" End If%>>锁定
			<input type="radio" name="istate" value="3" <%If rs("istate")=3 Then Response.Write " checked" End If%>>审核
		</td>
    </tr>
    <tr class="tdbg">
      <td>是否为推荐<%=oblog.CacheConfig(69)%>：</td>
      <td>
			<input type="radio" name="isbest" value="1" <%If rs("isbest")=1 Then Response.Write " checked" End If%>>推荐
			<input type="radio" name="isbest" value="0" <%If rs("isbest")=0 Then Response.Write " checked" End If%>>不推荐
		</td>
    </tr>
    <tr class="tdbg">
      <td colspan="2" align="center">
			<input type="submit" value=" 提交 " name="B1">
			<input type="reset" value=" 重置 " name="B2">
		</td>
    </tr>


  </form>
</table>
		</div>
	</div>
<%End Sub
Sub ReCountTeamInfo(teamid)
	Dim rst,c1,c2,c3,c4
	Set rst=oblog.execute("select Count(userid) From oblog_teamusers Where teamid=" & teamid)
	If not rs.Eof Then
		c1=OB_IIf(rst(0),0)
	Else
		c1=0
	End If
	Set rst=oblog.execute("select Count(postid) From oblog_teampost Where idepth=0 And teamid=" & teamid)
	If not rs.Eof Then
		c2=OB_IIf(rst(0),0)
	Else
		c2=0
	End If
	Set rst=oblog.execute("select Count(postid) From oblog_teampost Where idepth>0 And teamid=" & teamid)
	If not rs.Eof Then
		c3=OB_IIf(rst(0),0)
	Else
		c3=0
	End If
	oblog.execute "Update oblog_team Set iCount0=" & c1 & ",iCount1=" & c2 & ",iCount2=" & c3 & " Where teamid=" & teamid
	Set rst=Nothing
End Sub
sub modifyteaminfo()
	Dim name, rs, intro, sql, str,ico,tags,t1,t2,t3,t4,t5,team_domain,team_domainroot
	Dim CreaterName,ManagerMame,t_Name,ClassID,CreaterID,ManagerID,istate,isbest,istate_0
	Dim trs
    intro = Trim(Request.Form("intro"))
	ico = Trim(Request.Form("ico"))
    t1 = Trim(Request.Form("t1"))
    t2 = Trim(Request.Form("t2"))
    t3 = Trim(Request.Form("t3"))
    t4 = Trim(Request.Form("t4"))
    t5 = Trim(Request.Form("t5"))
    tags = Trim(Request.Form("tags"))
    team_domain = Trim(Request.Form("team_domain"))
    team_domainroot = Trim(Request.Form("team_domainroot"))
    CreaterName = Trim(Request.Form("creatername"))
    ManagerMame = Trim(Request.Form("managername"))
    t_Name = Trim(Request.Form("t_name"))
    ClassID = Int(Trim(Request.Form("ClassID")))
    istate = Int(Trim(Request.Form("istate")))
    istate_0 = Int(Trim(Request.Form("istate_0")))
    isbest = Trim(Request.Form("isbest"))
	If CreaterName = "" Or ManagerMame = "" Then
     	oblog.ShowMsg ("创建者或者管理员不能为空！"),""
        Exit Sub
	Else
		Set trs = oblog.Execute ("select userid FROM oblog_user WHERE username='"&CreaterName&"'")
		If trs.EOF Then
	     	oblog.ShowMsg ("创建者不存在！"),""
			Exit Sub
		Else
			CreaterID = trs(0)
		End If
		trs.Close
		Set trs = oblog.Execute ("select userid FROM oblog_user WHERE username='"&ManagerMame&"'")
		If trs.EOF Then
	     	oblog.ShowMsg ("创建者不存在！"),""
			Exit Sub
		Else
			ManagerID = trs(0)
		End If
		trs.Close
		Set trs = Nothing
	End If
	If t_Name = "" Then
     	oblog.ShowMsg ("名称不能为空！"),""
        Exit Sub
	Else
		t_Name=Left(t_Name,25)
	End if
	If t1="2"  Then
		If  t2="" Or Not isNumeric(t2) Then
			oblog.ShowMsg ("请输入加入时的积分限制"),""
	        Exit Sub
	     Else
	     	t2=Int(t2)
	     End If
	Else
		t2=0
	End If
	Set rs=Server.CreateObject("Adodb.Recordset")
    rs.Open "select * from oblog_team where teamid=" & teamid,conn,1,3
    If Not rs.EOF Then
		rs("t_Name") = t_Name
    	rs("t_ico")=ico
    	rs("joinlimit")=t1
    	rs("joinscores")=t2
		rs("viewlimit") = OB_IIF(t4,"-1")
		If t4 = "1" And t5<>"" Then rs("viewpassword")=MD5(t5)
    	rs("intro")=intro
    	rs("createrid")=CreaterID
    	rs("creatername")=CreaterName
    	rs("managerid")=ManagerID
    	rs("managername")=ManagerMame
		rs("classid")=classid
		rs("t_tags") = tags
		rs("istate") = istate
		rs("isbest") = OB_IIF(isbest,"0")
		If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" Then
			rs("t_domain")=team_domain
			rs("t_domainroot")=team_domainroot
		End if
    	rs.Update
    	str = "" & oblog.CacheConfig(69) & "信息修改完成"
    Else
    	str = "" & oblog.CacheConfig(69) & "信息不存在"
    End If
	'积分操作
	TeamScore teamid,istate_0
    rs.Close
    Set rs=Nothing
	WriteSysLog "进行了修改"&oblog.CacheConfig(69)&"操作，目标"&oblog.CacheConfig(69)&"ID："&teamid&"",oblog.NowUrl&"?"&Request.QueryString
    oblog.ShowMsg str, ""
End Sub
Set oblog = Nothing
%>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</body>
</html>
<script language="javascript">
function getImg(){
	if (document.oblogform.ico.value!=""){
		document.oblogform.imgIcon.src='<%=blogdir%>'+document.oblogform.ico.value;
	}
}
</script>