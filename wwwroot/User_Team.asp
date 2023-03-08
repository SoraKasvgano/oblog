<!--#include file="user_top.asp"-->
<%
Dim DivId,ContentId,rs,sql,sDisable
Dim teamId,teamname,managerid
DivId=Request("div")
If DivId="" Then DivId=11
DivId=Cint(DivId)
ContentId=Request("Content")
If ContentId="" Then ContentId=21
ContentId=Cint(ContentId)
'如果有TeamId,则取得该群组的基本信息
teamId=Request("teamid")
If teamId<>"" Then
	teamId=CLng(teamId)
	Set rs=oblog.Execute("select * From oblog_team Where istate>0 and teamid=" & teamId)
	teamname=rs("t_name")
	managerid=rs("managerid")
	Set rs=Nothing
	sDisable=""
Else
	sDisable=" disabled"
End If

%>
<script language="javascript">
function getImg(){
	if (document.oblogform.ico.value!=""){
		document.oblogform.imgIcon.src=document.oblogform.ico.value;
	}
}
function doMenu1(URL){
	document.getElementById("teamFrame").src=URL;
	document.getElementById("swin2").style.display = "block";
	}

//window.onload=function(){
//	var trs=user_team_left.getElementsByTagName("ol")
//	for(var i=0;i<trs.length;i++){
//		trs[i].style.backgroundColor=((i%2==0)?"#fff":"#F5FBFF");
//	}
//}
</script>

<%
'teamusers: state 1有效;2申请加入3被邀请4 副管理员 5 管理员
dim action,id
action=Request("action")
id=Trim(Request("id"))
select case action
	case "listjoinedteam"
		call listjoinedteam
	case "creatteam"
		call creatteam
	case "maketeam"
		call maketeam
	case "listuser"
		call listuser
	case "invite"
		call invite
	case "exitteam"
		call exitteam
	case "teamadmin"
		call teamadmin
	case "modifystate"
		call modifystate
	case "del"
		call del
	case "modifyteaminfo"
		call modifyteaminfo
	Case "members"
		Call ListMembers("",Cint(Request("cmd")))
	Case "state"
		Call MemberState
	case "links"
		call ShowAddon(1)
	case "showconfig"'选择群组模板
		call showconfig
	case "modiconfig"'修改群组主模板
		call modiconfig
	case "modiviceconfig"'修改群组副模板
		call modiviceconfig
	case "bakskin"'备份模板
		call bakskin
	case "announce"
		call ShowAddon(2)
	case "saveaddon"
		Call SaveAddon
	case "listmanageteam"
		call listmanageteam
	Case "teammanager"
		Call teammanager
	case else
		call main
		Response.Write("<div style='display:none'>")'输出一个闭合标签
end select
%>

</body>
</html>
<%
sub listmanageteam()
%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr2">
			<th>
				<table id="ListManageTeamTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"></td>
						<td class="t3"><%=oblog.CacheConfig(69)%></td>
						<td class="t4">帖子</td>
						<td class="t5">成员</td>
						<td class="t6">操作</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<table id="ListManageTeam" class="TableList" cellpadding="0">
<%
dim rs,i
set rs=oblog.execute("select a.* From oblog_team a,oblog_teamusers b Where b.userid=" & oblog.l_uid&" and a.teamid=b.teamid and b.state=5 and a.istate>0")
while not rs.eof
	i=i+1
%>
						<tr>
							<td class="t1">
								<%=i%>
							</td>
							<td class="t2">
								<img class="t_ico" src="<%=oblog.filt_html(ProIco(rs("t_ico"),3))%>">
							</td>
							<td class="t3">
								<%="<a href=""group.asp?gid="&rs("teamid")&""" target=""_blank"">"&rs("t_name")&"</a>"%>
								<div class="time">创建时间：<%=formatdatetime(rs("createtime"),0)%></div>
							</td>
							<td class="t4">
								<%=rs("icount1")%>
							</td>
							<td class="t5">
								<%=rs("icount0")%>
							</td>
							<td class="t6">
<%
select Case Cint(OB_IIF(rs("istate"),2))
Case 1
	Response.write "待审"
Case 2
	Response.write "锁定"
Case 3
%>
								<a href="user_team.asp?action=teamadmin&teamid=<%=rs("teamid")%>"><span class="green">管理</span></a>
<%
End select
%>
							</td>
						</tr>
<%
rs.movenext
wend
set rs=nothing
%>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/18.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
end sub

sub listjoinedteam()
dim rs,i
set rs=oblog.execute("select a.teamid,a.t_name,b.addtime,t_ico,icount0,b.post_all From oblog_team a,oblog_teamusers b Where a.teamid=b.teamid And b.userid=" & oblog.l_uid & " And b.state=3")
%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr2">
			<th>
				<table id="ListManageTeamTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"></td>
						<td class="t3"><%=oblog.CacheConfig(69)%></td>
						<td class="t4">帖子</td>
						<td class="t5">成员</td>
						<td class="t6">操作</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<table id="ListManageTeam" class="TableList" cellpadding="0">
<%
If rs.Eof Then
%>
						<tr>
							<td class="t1"></td>
							<td class="t2"></td>
							<td class="t3"><div class="msg">未加入他人创建的<%=oblog.CacheConfig(69)%></div></td>
							<td class="t4"></td>
							<td class="t5"></td>
							<td class="t6"></td>
						</tr>
<%
Else
Do while not rs.eof
i=i+1
%>
						<tr>
							<td class="t1">
								<%=i%>
							</td>
							<td class="t2">
								<img class="t_ico" src="<%=oblog.filt_html(ProIco(rs("t_ico"),3))%>">
							</td>
							<td class="t3">
								<a href="group.asp?gid=<%=rs("teamid")%>" target="_blank"><%=rs("t_name")%></a>
								<div class="time">加入时间：<%=FormatDateTime(OB_IIF(rs("addtime"),Now()),0)%></div>
							</td>
							<td class="t4">
								<%=OB_IIF(rs("post_all"),0)%>
							</td>
							<td class="t5">
								<%=OB_IIF(rs("icount0"),0)%>
							</td>
							<td class="t6">
								<a href="user_team.asp?action=exitteam&teamid=<%=rs("teamid")%>" onclick="if (confirm('确认要退出该<%=oblog.CacheConfig(69)%>吗?')==false) return false;"><span class="red">退出</span></a>
							</td>
						</tr>
<%
rs.movenext
Loop
End If
set rs=nothing
%>
					</table>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/18.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
end sub

Sub creatteam()      '创建圈子页面部分
	Dim rs
	If oblog.l_Group(16,0)=0 Then
		oblog.AddErrStr ("您目前所属的等级目前不允许创建新的" &oblog.CacheConfig(69))
	    oblog.showUserErr
	    Response.End
	End if
	'检查之前是够有待审核的申请
	Set rs=oblog.Execute("select t_name From oblog_team Where istate=1 And managerid=" & oblog.l_uid)
	If Not rs.Eof Then
		oblog.adderrstr("您之前创建的 " & rs(0) & " 还没有被审核通过，不能再创建新的" &oblog.CacheConfig(69))
		oblog.showusererr
		rs.Close
	End If
	'检查目前管理的总数
	Set rs=oblog.Execute("select count(teamid) From oblog_team Where  managerid=" & oblog.l_uid)
	If rs(0)>=oblog.l_Group(16,0) Then
		oblog.adderrstr("您目前已管理 " & rs(0) & " 个" &oblog.CacheConfig(69)& "，达到系统的限额。" )
		oblog.showusererr
		rs.Close
	End If
	Set rs=Nothing
%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<form name="oblogform" id="NewTeam" method="post" action="user_team.asp?action=maketeam" >
					<table cellpadding="0">
						<tr>
							<td class="title">
								<label for="name"><%=oblog.CacheConfig(69)%>名称：</label>
							</td>
							<td>
								<input type="text" name="name" id="name" size="30" />
							</td>
						</tr>
<%If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" Then%>
						<tr>
							<td class="title">
								<label for="team_domain"><%=oblog.CacheConfig(69)%>域名：</label></td>
							<td>
								<input name="team_domain" id="team_domain" type="text" value="" size=10 maxlength=20 />
								<select name="team_domainroot" >
									<%=oblog.type_domainroot("",1)%>
								</select>
							</td>
						</tr>
						<tr <%If CBool(true_domain) Then response.write "style=""display:none;"""%>>
							<td class="title">
								隐藏转向URL：
							</td>
							<td>
								<label><input type="radio" value="1" name="hideurl" />是</label> &nbsp;&nbsp;
								<label><input type="radio" value="0" name="hideurl" checked />否</label>
							</td>
						</tr>
<%End if%>
						<tr>
							<td class="title">
								<%=oblog.CacheConfig(69)%>标记图片：
							</td>
							<td>
								<div class="user_face">
									<img id = "imgIcon" src="images/default_groupico.gif" class="t_ico">
									<p><iframe id="d_file" frameborder="0" src="upload.asp?tMode=8&re=&teamid=<%=teamId%>" width="300" height="80" scrolling="no"></iframe><br/>
									<label>外部调用：<input name="ico" id="ico" type="text" value="images/default_groupico.gif" size="50" maxlength="200" onblur="getImg();" /></label><br/>
									你可以直接输入一个有效的图片地址,也可以在这里直接选择一个系统可用的图片</p>
								</div>
							</td>
						</tr>
						<tr>
							<td class="title">
								<label for="classid"><%=oblog.CacheConfig(69)%>类别：</label>
							</td>
							<td>
								<select name="classid" id="classid" >
									<%=oblog.show_class("log",0,2)%>
								</select>
							</td>
						</tr>
						<tr>
							<td class="title">
								<label for="tags"><%=oblog.CacheConfig(69)%>标签：</label>
							</td>
							<td>
								<input type="text" name="tags" id="tags" size="50" />(最多支持5个，以逗号间隔)
							</td>
						</tr>
						<tr>
							<td class="title">
								<%=oblog.CacheConfig(69)%>加入条件：
							</td>
							<td>
								<label><input type="radio" name="t1" value="-1" />任意加入</label><br />
								<label><input type="radio" name="t1" value="0" checked />申请加入</label><br />
								<label><input type="radio" name="t1" value="1" />仅可邀请</label><br />
								<label for="t2"><input type="radio" name="t1" value="2" />积分限制，需大于</label><input type="text" name="t2" id="t2" size="5" maxlength="8" /><label for="t2">积分才能申请</label>
							</td>
						</tr>
						<tr>
							<td class="title">
								<%=oblog.CacheConfig(69)%>访问权限：
							</td>
							<td>
								<label><input type="radio" name="t4" value="-1" checked />无限制</label><br />
								<label><input type="radio" name="t4" value="0" /><%=oblog.CacheConfig(69)%>成员可见</label><br />
								<label><input type="radio" name="t4" value="1" />凭密码访问，密码<input type=text name="t5" size="20" maxlength="20" /></label>
							</td>
						</tr>
						<tr>
							<td class="title">
								允许非成员参与回复：
							</td>
							<td>
								<label><input type="radio" name="t3" value="1" checked />是</label>&nbsp;&nbsp;
								<label><input type="radio" name="t3" value="0" />否</label>（须登录状态才可回复）
							</td>
						</tr>
						<tr>
							<td class="title">
								<label for="intro"><%=oblog.CacheConfig(69)%>申请说明：</label>
							</td>
							<td>
								<textarea rows="5" name="intro" id="intro" cols="45"></textarea>
							</td>
						</tr>
						<tr>
							<td></td>
							<td>
								<input type="submit" id="Submit" value=" 确 认 提 交 " name="b1" />
							</td>
						</tr>
					</table>

				</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
End sub

Sub MakeTeam()
	Dim rs
	Set rs=Server.CreateObject("Adodb.Recordset")
    If oblog.l_Group(16,0)=0 Then
		oblog.AddErrStr ("您目前所属的等级目前不允许创建新的" &oblog.CacheConfig(69))
	    oblog.showUserErr
	    Response.End
	End If
	If oblog.l_uScores<CLng(oblog.CacheScores(11)) Then
		oblog.AddErrStr ("您的积分不足，无法申请" &oblog.CacheConfig(69))
	    oblog.showUserErr
	    Response.End
	End if
	'检查之前是够有待审核的申请
	Set rs=oblog.Execute("select t_name From oblog_team Where istate=1 And managerid=" & oblog.l_uid)
	If Not rs.Eof Then
		oblog.adderrstr("您之前创建的 " & rs(0) & " 还没有被申请通过，不能再创建新的" &oblog.CacheConfig(69))
		oblog.showusererr
		rs.Close
	End If
	'检查目前管理的总数
	Set rs=oblog.Execute("select count(teamid) From oblog_team Where  managerid=" & oblog.l_uid)
	If rs(0)>=oblog.l_Group(16,0) Then
		oblog.adderrstr("您目前已管理 " & rs(0) & " 个" &oblog.CacheConfig(69)& "，达到系统的限额。" )
		rs.Close
		oblog.showusererr
	End If
	rs.Close
    Dim name, intro, sql, teamid, str,ico,tags,t1,t2,t3,tid,classid,t4,t5,team_domain,team_domainroot,hideurl
    name = oblog.filt_badword(Trim(Request.Form("name")))
	name = oblog.filt_badstr(name)
    intro = Trim(Request.Form("intro"))
	ico = Trim(Request.Form("ico"))
    t1 = Trim(Request.Form("t1"))
    t2 = Trim(Request.Form("t2"))
    t3 = Trim(Request.Form("t3"))
    t4 = Trim(Request.Form("t4"))
    t5 = Trim(Request.Form("t5"))
    tags = Trim(Request.Form("tags"))
	classid = Trim(Request.Form("classid"))
	team_domain = LCase(Trim(Request.Form("team_domain")))
	team_domainroot = Trim(Request.Form("team_domainroot"))
	hideurl = Trim(Request.Form("hideurl"))
	if classid="" Or classid = 0 then
		oblog.ShowMsg "" &oblog.CacheConfig(69)& "分类不能为空！","back"
		Exit Sub
	else
		classid=CLng(classid)
	end if
    If name="" Then
    	oblog.ShowMsg "名称不能为空！","back"
		exit sub
    Else
    	name=Left(name,25)
	End If
	If intro="" Then
    	oblog.ShowMsg "申请说明不能为空！","back"
		exit sub
    Else
    	intro=Left(intro,240)
	End If
	If t1="2"  Then
		If  t2="" Or Not isNumeric(t2) Then
			oblog.ShowMsg "请输入加入时的积分限制","back"
			exit sub
	     Else
	     	t2=CLng(t2)
	     End If
	Else
		t2=0
	End If

	If t4="1"  Then
		If  t5="" Then
			oblog.ShowMsg "请输入访问密码","back"
			exit sub
	     End If
	Else
		t5=""
	End If
	If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" Then
		If team_domain = "" Or oblog.strLength(team_domain) > 10 Then oblog.adderrstr  ("域名不能为空(不能大于10个字符)！")
		If team_domain <> Request("old_userdomain") And oblog.strLength(team_domain) < 4 Then oblog.adderrstr  ("域名不能小于4个字符！")
		If oblog.chk_regname(team_domain) Then oblog.adderrstr  ("此域名系统不允许注册！")
		If oblog.chk_badword(team_domain) > 0 Then oblog.adderrstr  ("域名中含有系统不允许的字符！")
		If oblog.chkdomain(team_domain) = False Then oblog.adderrstr  ("域名不合规范，只能使用小写字母，数字！")
		If team_domainroot = "" Then oblog.adderrstr  ("域名根不能为空！")
		If oblog.CheckDomainRoot(team_domainroot,1) = False Then oblog.adderrstr  ("域名根不合法！")
	End If
	If team_domain="" Then
	rs.Open "select * from oblog_team where t_name='" & name & "' " ,conn,1,3
	Else
    rs.Open "select * from oblog_team where t_name='" & name & "' or t_domain='"&team_domain&"'" ,conn,1,3
	End If
    If Not rs.EOF Then
        Set rs = Nothing
    	oblog.ShowMsg "此" & oblog.CacheConfig(69) & "名 或域名已经存在！","back"
		exit sub
    Else
    	rs.AddNew
    	rs("t_name")=name
    	rs("t_ico")=ico
    	rs("joinlimit")=t1
    	rs("joinscores")=t2
		rs("otherpost")=t3
    	rs("otherpost")=0
    	rs("createrid")=oblog.l_uid
    	rs("creatername")=oblog.l_uname
    	rs("managerid")=oblog.l_uid
    	rs("managername")=oblog.l_uname
    	rs("createtime")=oblog.ServerDate(Now)
    	If oblog.CacheConfig(49) = "1" Then
			rs("istate")=1
		Else
			rs("istate")=3
		End if
    	rs("icount0")=1
    	rs("intro")=intro
		rs("classid")=classid
		rs("t_tags")=tags
		rs("viewlimit")=t4
		If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" Then
			rs("t_domain")=team_domain
			rs("t_domainroot")=team_domainroot
			rs("hideurl") = hideurl
		End if
		If t4 = 1 Then rs("viewpassword")=MD5(t5)
    	rs.Update
    	rs.Close
    	rs.Open "select Max(teamid) From oblog_team Where createrid=" & oblog.l_uid
    	tid=rs(0)
    	rs.Close
    	oblog.Execute "Insert into oblog_teamusers(teamid,userid,state) values (" & tid & "," & oblog.l_uid & ",5)"
		oblog.GiveScore "" ,-1*Abs(oblog.CacheScores(11)),""
    	str = "" & oblog.CacheConfig(69) & ":" & name &"提交完成"
		If oblog.CacheConfig(49) = 1 Then str=str & "正在等待管理员审核"
    	oblog.ShowMsg str, "user_team.asp"
    End If
End Sub


Sub listuser()
	Dim grade,i
	sql="select state from oblog_teamusers where userid="&oblog.l_uid&" and teamid="&teamid&" and state=5"
	set rs=oblog.execute(sql)
	if rs.eof or rs.bof then
		set rs=nothing
		oblog.adderrstr("您的权限不足,操作无法完成！")
		oblog.showusererr
	end if
%>
<table class="TeamContent" cellpadding="0">
	<tr>
		<td class="t1">
			<div id="chk_idAll">
				<form method="post" action="user_team.asp?action=invite&teamid=<%=teamid%>">
				<ul id="UserMenu">
					<li><a href="#" onClick="return doMenu('swin1');">邀请用户</a></li>
					<li><a href="#" onClick="return doMenu('swin2');"><span class="red">转让<%=oblog.CacheConfig(69)%></span></a></li>
				</ul>
					<div id="swin1" style="display:none;position:absolute;top:41px;left:10px;z-index:100;">
						<form name="form1" method="post" action="user_friendurl.asp?action=addurl&t=<%=t%>">
						<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
							<tr>
								<td colspan='2' align='center' class='win_table_top'>邀请新用户加入我的<%=oblog.CacheConfig(69)%></td>
							</tr>
							<tr>
								<td class='win_table_td'><label for="T1">用户名：</label></td>
								<td><input type="text" name="T1" id="T1"  size="20"><input type="submit" id="Submit" name="b3" value="邀请" /></td>
							</tr>
							<tr>
								<td colspan='2' class="win_table_end"><input type="button" onClick="return doMenu('swin1');" value="关闭" title="关闭" /></td>
							</tr>
						</table>
						</form>
					</div>
					<div id="swin2" style="display:none;position:absolute;top:41px;left:10px;z-index:100;">
<% If oblog.l_uid=managerid Then%>
						<form method="POST" action="user_team.asp?action=teammanager&teamid=<%=teamid%>">
						<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
							<tr>
								<td colspan='2' align='center' class='win_table_top'>转让<%=oblog.CacheConfig(69)%></td>
							</tr>
							<tr>
								<td class='win_table_td'><label for="T1">用户名：</label></td>
								<td><input type="text" name="T1" id="T1" size="20"><input type="submit" id="Submit" name="b3" value="决定转让" /></td>
							</tr>
							<tr>
								<td colspan='2' class="win_table_end"><input type="button" onClick="return doMenu('swin2');" value="关闭" title="关闭" /></td>
							</tr>
						</table>
						</form>
<%End if%>
					</div>
					<div id="swin3"></div>
					<div id="swin4"></div>
					<div id="swin5"></div>
					<iframe id="DivShim" scrolling="no" frameborder="0" style="position:absolute;top:0px; left:0px;display:none"></iframe>
				</form>
				<table id="TeamListUserTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2">用户名称</td>
						<td class="t3">用户等级</td>
						<td class="t4">操作</td>
					</tr>
				</table>
				<table id="TeamListUser" cellpadding="0">
<%
	set rs=oblog.execute("select a.username,b.state,b.userid,b.id from oblog_teamusers b,oblog_user a where a.userid=b.userid and b.teamid="&teamid&" and b.state<>4")
	while not rs.eof
		select case cint(rs(1))
			case 3
				grade="普通用户"
			case 1
				grade="受邀但尚未回应"
			case 2
				grade="申请加入"
			case 5
				grade="管理员"
		end select
		If rs(2)=managerid Then grade = grade & "<font color=red>（群主）</font>"
		i=i+1
%>
					<tr>
						<td class="t1">
							<%=i%>
						</td>
						<td class="t2">
							<a href='blog.asp?name=<%=rs(0)%>' target="_blank"><%=rs(0)%></a>
						</td>
						<td class="t3">
							<%=grade%>
						</td>
						<td class="t4">
<%select case cint(rs(1))
case 3
%>
							<a href="user_team.asp?action=del&state=1&userid=<%=rs(2)%>&teamid=<%=teamid%>"><span class="red">删除</span></a>&nbsp; <a href="user_team.asp?action=modifystate&g1=3&g2=5&userid=<%=rs(2)%>&teamid=<%=teamid%>"><span class="blue">升为管理员</span></a>
<%  case 1%>
							<a href="user_team.asp?action=del&state=1&userid=<%=rs(2)%>&teamid=<%=teamid%>"><span class="red">删除邀请</span></a>&nbsp;
<%  case 2%>
							<a href="javascript:openScript('user_pm.asp?action=readteam&id=<%=rs(3)%>',450,380)"><span class="blue">查看申请</span></a>&nbsp;<a href="user_team.asp?action=modifystate&g1=2&g2=3&userid=<%=rs(2)%>&teamid=<%=teamid%>"><span class="green">同意申请</span></a>&nbsp;<a href="user_team.asp?action=del&state=2&userid=<%=rs(2)%>&teamid=<%=teamid%>"><span class="red">拒绝</span></a>
<%  case 5%>
							<a href="user_team.asp?action=modifystate&g1=5&g2=3&userid=<%=rs(2)%>&teamid=<%=teamid%>"><span class="red">降为普通用户</span></a>&nbsp;
							<a href="user_team.asp?action=del&state=5&userid=<%=rs(2)%>&teamid=<%=teamid%>"><span class="red">删除</span></a>
<%end select%>
						</td>
					</tr>

<%
	rs.movenext
wend
rs.close
%>
				</table>
			</div>
			<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
		</td>
<%
	teamadmin_top()
end sub


Sub ExitTeam()
    Dim sql, rs, str, name
    sql = "select t_name from oblog_team where teamid=" & teamid
    Set rs = oblog.Execute(sql)
    name = rs(0)
    rs.Close
    sql = "select userid from oblog_teamusers where teamid=" & teamid & " and userid=" & oblog.l_uid & " and state=5"
    rs.open sql, conn, 1, 1
    If rs.recordcount = 1 Then
        str = "您是" & oblog.CacheConfig(69) & name & "的管理员,无法退出该" & oblog.CacheConfig(69) & ",若要退出请先转移管理员权限"
        oblog.adderrstr (str)
        oblog.showusererr
    End If
    sql = "delete from oblog_teamusers where (teamid=" & teamid & " and userid=" & oblog.l_uid & " and state=3) or (teamid=" & teamid & " and userid=" & oblog.l_uid & " and state=5)"
    oblog.Execute (sql)
    str = "成功退出了" & oblog.CacheConfig(69) & ":" & name & ",您已不再是" & name & "的正式成员"
    oblog.ShowMsg str, ""
End Sub

sub teamadmin()
%>
				<!--管理<%=oblog.CacheConfig(69)%>-->
<%
	dim i,grade
	set rs=oblog.execute("select state from oblog_teamusers where userid="&oblog.l_uid&" and teamid="&teamid)
	if not rs.eof then
		if rs(0)<>5 then oblog.adderrstr ("无权限")
	else
		oblog.adderrstr ("无权限")
	end if
	if oblog.errstr<>"" then
		set rs=nothing
		oblog.showusererr
		exit sub
	end if
	set rs=oblog.execute("select * from oblog_team where teamid="&teamid)
	ReCountTeamInfo(teamid)

%>
<table class="TeamContent" cellpadding="0">
	<tr>
		<td class="t1">
			<div id="chk_idAll">
				<form name="oblogform" method="post" action="user_team.asp?action=modifyteaminfo&teamid=<%=teamid%>">
				<table class="TeamInfo" cellpadding="0">
					<tr>
						<td class="title">
							<label for="name"><%=oblog.CacheConfig(69)%>名称：</label>
						</td>
						<td>
							<input type="text" name="name" id="name" size="30" value="<%=rs("t_name")%>" disabled>
						</td>
					</tr>
<%If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" Then%>
					<tr>
						<td class="title">
							<label for="team_domain"><%=oblog.CacheConfig(69)%>域名：</label>
						</td>
						<td>
							<input name="team_domain" id="team_domain" type="text" value="<%=rs("t_domain")%>" size=10 maxlength=20 /> <select name="team_domainroot" ><%=oblog.type_domainroot(rs("t_domainroot"),1)%></select>
						</td>
					</tr>
					<tr>
						<td class="title">
							是否隐藏转向URL：
						</td>
						<td>
							<label><input type="radio" value="1" name="hideurl" <%If rs("hideurl")=1 Then Response.Write " checked" End If%> />是</label> &nbsp;&nbsp;
							<label><input type="radio" value="0" name="hideurl" <%If rs("hideurl")=0 Then Response.Write " checked" End If%>  />否</label>
						</td>
					</tr>
<%End if%>
					<tr>
						<td class="title">
							标记图片：
						</td>
						<td>
							<img class="t_ico" id="imgIcon" src="<%=oblog.filt_html(ProIco(rs("t_ico"),3))%>">
							<p><iframe id="d_file" frameborder="0" src="upload.asp?tMode=8&re=&teamid=<%=teamId%>" width="380" height="30" scrolling="no"></iframe><br/>
							<span class="blue">只支持jpg,gif,png,小于200k,默认尺寸为120*90</span><br/>
							<label>图片地址：<input name="ico"  type="text" value="<%=oblog.filt_html(rs("t_ico"))%>" size="50" maxlength="200" onblur="getImg();"></label>
							<br/><span class="blue">你可以直接输入一个有效的图片地址，也可以在这里直接选择一个系统可用的图片</span></p>
						</td>
					</tr>
					<tr>
						<td class="title">
							<label for="tags"><%=oblog.CacheConfig(69)%>标签：</label>
						</td>
						<td>
							<input type="text" name="tags" id="tags" size="50" value="<%=rs("t_tags")%>">（最多支持5个，以逗号间隔）
						</td>
					</tr>
					<tr>
						<td class="title">
							<%=oblog.CacheConfig(69)%>加入条件：
						</td>
						<td>
							<label><input type="radio" name="t1" value="-1" <%If rs("joinlimit")=-1 Then Response.Write " checked" End If%>>任意加入</label><br />
							<label><input type="radio" name="t1" value="0" <%If rs("joinlimit")=0 Then Response.Write " checked" End If%>>申请加入</label><br />
							<label><input type="radio" name="t1" value="1" <%If rs("joinlimit")=1 Then Response.Write " checked" End If%>>仅可邀请</label><br/>
							<label><input type="radio" name="t1" value="2" <%If rs("joinlimit")=2 Then Response.Write " checked" End If%>>积分限制，需大于<input type=text name="t2" size=5 maxlength=8 value="<%=rs("joinscores")%>">积分才能申请</label>
						</td>
					</tr>
					<tr>
						<td class="title">
							<%=oblog.CacheConfig(69)%>访问权限：
						</td>
						<td>
							<label><input type="radio" name="t4" value="-1" <%If rs("viewlimit")=-1 Then Response.Write " checked" End If%>>无限制</label><br />
							<label><input type="radio" name="t4" value="0" <%If rs("viewlimit")=0 Then Response.Write " checked" End If%>><%=oblog.CacheConfig(69)%>成员可见</label><br />
							<label><input type="radio" name="t4" value="1" <%If rs("viewlimit")=1 Then Response.Write " checked" End If%>>凭密码访问，密码<input type=text name="t5" size=20 maxlength=20 value="">（不修改请留空）</label>
						</td>
					</tr>
					<tr>
						<td class="title">
							允许非成员回复：
						</td>
						<td>
							<label><input type="radio" name="t3" value="1" <%If rs("otherpost")=1 Then Response.Write " checked" End If%>>是</label>&nbsp;&nbsp;
							<label><input type="radio" name="t3" value="0" <%If rs("otherpost")=0 Then Response.Write " checked" End If%>>否</label>
						</td>
					</tr>
					<tr>
						<td class="title">
							<%=oblog.CacheConfig(69)%>简介：
						</td>
						<td>
							<textarea rows="5" name="intro" cols="45"><%=rs("intro")%></textarea>
						</td>
					</tr>
					<tr>
						<td class="title"></td>
						<td>
							<input type="submit" id="Submit" value="确认保存" name="B1">
						</td>
					</tr>
				</table>
				</form>
			</div>
			<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
		</td>
<%
	teamadmin_top()
end sub

sub modifystate()

	dim g1,g2,sql,rs,userid,teamid,str
	g1=CLng(Request.QueryString("g1"))
	g2=CLng(Request.QueryString("g2"))  'g1是原始状态,G2是目标状态
	userid=CLng(Request.QueryString("userid"))
	teamid=CLng(Request.QueryString("teamid"))
	sql="select state from oblog_teamusers where userid="&oblog.l_uid&" and teamid="&teamid
	set rs=oblog.execute(sql)

	If userid=managerid Then
		oblog.adderrstr("您无权对群主进行操作！")
		oblog.showusererr
		Exit Sub
	end If

	if cint(rs(0))<>5 then
		oblog.adderrstr("您的权限不足,操作无法完成！")
		oblog.showusererr
	end if
	rs.close
	if g1=5 then
		sql="select state from oblog_teamusers where state=5 and teamid="&teamid&""
		rs.open sql,conn,1,1
		if rs.recordcount=1 then
			str="该管理员是"&oblog.CacheConfig(69)&"中唯一的管理员,无法降级"
			oblog.adderrstr(str)
			oblog.showusererr
		end if
	end if

	set rs=nothing

	sql="update oblog_teamusers set state="&g2&" where userid="&userid&" and state="&g1&" and teamid="&teamid
	oblog.execute(sql)
	oblog.ShowMsg "用户状态修改成功",""
end sub

Sub del()
	dim state,teamid,userid,sql,rs,str

	state=CLng(Request("state"))
	teamid=CLng(Request("teamid"))
	userid=CLng(Request("userid"))

	sql="select state from oblog_teamusers where userid="&oblog.l_uid&" and teamid="&teamid
	set rs=oblog.execute(sql)

	If userid=managerid Then
		oblog.adderrstr("您无权对群主进行操作！")
		oblog.showusererr
		Exit Sub
	end if

	if cint(rs(0))<>5 then
		oblog.adderrstr("您的权限不足,操作无法完成！")
		set rs=nothing
		oblog.showusererr
		exit sub
	end if
	set rs=Server.CreateObject("adodb.recordset")
	if state=5 then
		sql="select state from oblog_teamusers where state=5 and teamid="&teamid
		rs.open sql,conn,1,1
		if rs.recordcount=1 then
			str="该管理员是"&oblog.CacheConfig(69)&"中唯一的管理员,无法删除"
			oblog.adderrstr(str)
			oblog.showusererr
			exit sub
		end if
	end if

	sql="delete from oblog_teamusers where teamid="&teamid&" and userid="&userid
	oblog.execute(sql)
	oblog.ShowMsg "成功删除相关信息",""
end sub

sub modifyteaminfo()
	Dim name, rs, intro, sql, str,ico,tags,t1,t2,t3,t4,t5,team_domain,team_domainroot,hideurl
    intro = Trim(Request.Form("intro"))
	ico = Trim(Request.Form("ico"))
    t1 = Trim(Request.Form("t1"))
    t2 = Trim(Request.Form("t2"))
    t3 = Trim(Request.Form("t3"))
    t4 = Trim(Request.Form("t4"))
    t5 = Trim(Request.Form("t5"))
    tags = Trim(Request.Form("tags"))
    team_domain = LCase(Trim(Request.Form("team_domain")))
    team_domainroot = Trim(Request.Form("team_domainroot"))
    hideurl = Trim(Request.Form("hideurl"))
	If intro="" Then
    	oblog.adderrstr ("介绍不能为空！")
        oblog.showusererr
    Else
    	intro=Left(intro,240)
	End If
	If t1="2"  Then
		If  t2="" Or Not isNumeric(t2) Then
			oblog.adderrstr ("请输入加入时的积分限制")
	        oblog.showusererr
	     Else
	     	t2=CLng(t2)
	     End If
	Else
		t2=0
	End If
	If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" Then
		If team_domain = "" Or oblog.strLength(team_domain) > 10 Then oblog.adderrstr  ("域名不能为空(不能大于10个字符)！")
		If team_domain <> Request("old_userdomain") And oblog.strLength(team_domain) < 4 Then oblog.adderrstr  ("域名不能小于4个字符！")
		If oblog.chk_regname(team_domain) Then oblog.adderrstr  ("此域名系统不允许注册！")
		If oblog.chk_badword(team_domain) > 0 Then oblog.adderrstr  ("域名中含有系统不允许的字符！")
		If oblog.chkdomain(team_domain) = False Then oblog.adderrstr  ("域名不合规范，只能使用小写字母，数字！")
		If team_domainroot = "" Then oblog.adderrstr  ("域名根不能为空！")
		If oblog.CheckDomainRoot(team_domainroot,1) = False Then oblog.adderrstr  ("域名根不合法！")
	End if
	If oblog.ErrStr <> "" Then oblog.showusererr

Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open "select * from oblog_team where t_domain='"&team_domain&"' and t_domainroot='"&team_domainroot&"' and not (teamid=" & teamid & " And managerid=" & oblog.l_uid&") ",conn,1,1
	If Not rs.eof  Then
		str = "" & oblog.CacheConfig(69) & "域名冲突，请更换域名"
		rs.Close
		oblog.ShowMsg str, "back"
		Exit Sub
	End If
	Set rs=Nothing
	Set rs=Server.CreateObject("Adodb.Recordset")
    rs.Open "select * from oblog_team where teamid=" & teamid & " And managerid=" & oblog.l_uid,conn,1,3
    If Not rs.EOF Then
    	rs("t_ico")=ico
    	rs("joinlimit")=OB_IIF(t1,0)
    	rs("joinscores")=OB_IIF(t2,0)
		rs("otherpost")=OB_IIF(t3,0)
		rs("viewlimit") = OB_IIF(t4,0)
		If t4 = 1 And t5<>"" Then rs("viewpassword")=MD5(t5)
    	rs("intro")=intro
		rs("t_tags") = tags

		If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" Then
			rs("t_domain")=Left(team_domain,10)
			rs("t_domainroot")=team_domainroot
			rs("hideurl") = hideurl
		End if
    	rs.Update
    	str = "" & oblog.CacheConfig(69) & "信息修改完成"
    Else
    	str = "" & oblog.CacheConfig(69) & "信息不存在"
    End If
    rs.Close
    Set rs=Nothing
    oblog.ShowMsg str, ""
End sub

sub invite()
	dim username,rs,id,teamid,sql,str
	username=oblog.filt_badstr(Trim(Request.Form("t1")))
	teamid=CLng(Request.QueryString("teamid"))
	sql="select userid from oblog_user where username='"&username&"'"
	set rs=oblog.execute(sql)
	If rs.Eof Then
		str="用户名"&username&"不存在"
	Else
		id=CLng(rs(0))
		set rs=oblog.execute("select * From oblog_teamusers Where teamid=" & teamid & " And userid=" & id)

		if rs.eof then
			sql="insert into oblog_teamusers (teamid,userid,state) values("&teamid&","&id&",1)"
			oblog.execute(sql)
			str="成功向"&username&"发出邀请"
		else
			select Case rs("state")
				Case 3
					str="此用户已经是" &oblog.CacheConfig(69)& "的成员"
				Case 1
					str="此用户已经被邀请了"
				Case 2
					str="此用户已经发出申请，请通过审核即可"
				Case 5
					str="此用户是" &oblog.CacheConfig(69)& "管理员,不需要进行申请"
			End select
		End If
	end if
	Set rs=Nothing
	oblog.ShowMsg str,""
end sub

'teamusers: state 1被邀请2申请加入3成员4 副管理员 5 管理员
Sub ListMembers(teamid,cmd)
	Dim sTitle,i,SqlPart,grade
	i=0
	If teamid<>"" Then SqlPart="And b.teamid=" & CLng(teamid) & " "
	select Case cmd
		Case 1
			sTitle="我发出的" & oblog.CacheConfig(69) & "邀请"
			Sql="select a.userid,c.username,a.addtime,b.teamid,b.t_name,a.state From oblog_teamusers a,oblog_team b,oblog_user c  Where a.teamid=b.teamid And state=1 And a.userid=c.userid  And a.userid<>" & oblog.l_uid &" And b.managerid=" & oblog.l_uid & SqlPart
		Case 2
			sTitle="我接收到的" & oblog.CacheConfig(69) & "邀请"
			Sql="select a.userid,c.username,a.addtime,b.teamid,b.t_name,a.state From oblog_teamusers a,oblog_team b,oblog_user c  Where a.teamid=b.teamid And state=1 And a.userid=c.userid  And a.userid=" & oblog.l_uid &" And b.managerid<>" & oblog.l_uid & SqlPart
		Case 3
			sTitle="我发出的" & oblog.CacheConfig(69) & "申请"
			Sql="select a.userid,c.username,a.addtime,b.teamid,b.t_name,a.state From oblog_teamusers a,oblog_team b,oblog_user c  Where a.teamid=b.teamid And state=2 And a.userid=c.userid  And a.userid=" & oblog.l_uid &" And b.managerid<>" & oblog.l_uid & SqlPart
		Case 4
			sTitle="我接收到的" & oblog.CacheConfig(69) & "申请"
			Sql="select a.userid,c.username,a.addtime,b.teamid,b.t_name,a.state,a.info From oblog_teamusers a,oblog_team b,oblog_user c  Where a.teamid=b.teamid And state=2 And a.userid=c.userid  And a.userid<>" & oblog.l_uid &" And b.managerid=" & oblog.l_uid & SqlPart
	End select
	'Response.Write Sql
	Set rs=oblog.execute(Sql)

%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr2">
			<th>
				<table id="ListMembersTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"><%=sTitle%></td>
						<td class="t3">状态</td>
						<td class="t4">操作</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<table id="ListMembers" class="TableList" cellpadding="0">
<%
If rs.Eof Then
%>
						<tr>
							<td class="ListMembersTdMsg" colspan="4"><div class="msg">没有相关数据</div></td>
						</tr>
<%
Else
	while not rs.eof
	select case cint(rs("state"))
		case 3
			grade="普通用户"
		case 2
			grade="申请待批准"
		case 1
			grade="受邀但尚未回应"
		case 5
			grade="管理员"
	end select
	i=i+1
%>
						<tr>
							<td class="t1">
								<%=i%>
							</td>
							<td class="t2">
								<a href="group.asp?gid=<%=rs("teamid")%>" target="_blank"><%=rs("t_name")%></a>
								<div><a href="go.asp?userid=<%=rs("userid")%>" target="_blank"><%=rs("username")%></a><span class="time">posted <%=rs("addtime")%></span></div>
							</td>
							<td class="t3">
								<%=grade%>
							</td>
							<td class="t4">
<%
select Case cmd
Case 1
%>
								<a href="user_team.asp?action=state&cmd=1&state=0&userid=<%=rs("userid")%>&teamid=<%=rs("teamid")%>">取消邀请</a>
<%
Case 2
%>
								<a href="user_team.asp?action=state&cmd=2&state=3&userid=<%=rs("userid")%>&teamid=<%=rs("teamid")%>">接受</a>&nbsp;
								<a href="user_team.asp?action=state&cmd=2&state=0&userid=<%=rs("userid")%>&teamid=<%=rs("teamid")%>">拒绝</a>
<%
Case 3
%>
								<a href="user_team.asp?action=state&cmd=3&state=0&userid=<%=rs("userid")%>&teamid=<%=rs("teamid")%>">取消申请</a>
<%
Case 4
%>
								<a href="user_team.asp?action=state&cmd=4&state=3&userid=<%=rs("userid")%>&teamid=<%=rs("teamid")%>" title="<%=oblog.filt_html(rs("info"))%>">接受</a>&nbsp;
								<a href="user_team.asp?action=state&cmd=4&state=0&userid=<%=rs("userid")%>&teamid=<%=rs("teamid")%>" title="<%=oblog.filt_html(rs("info"))%>" >拒绝</a>
<%
End select
%>
							</td>
						</tr>
<%
		rs.movenext
	wend
End If
rs.close
%>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/18.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
End sub

Sub MemberState()
	Dim userId,sMsg,iState
	userid=clng(Request("userid"))
	iState=Cint(Request("state"))
	select Case Int(Request("cmd"))
		Case 1
			'取消邀请(由管理员操作)
			Set rs=oblog.execute("select teamid From oblog_team Where teamid=" & teamid & " And managerid=" & oblog.l_uid)
			If Not rs.Eof Then
				oblog.execute "Delete From oblog_teamusers Where userid=" & userid & " And teamid=" & teamId
				sMsg="已取消对该用户的邀请"
			Else
				sMsg="您无权取消对该用户的邀请"
			End If
			Set rs=Nothing
		Case 2
			'接受邀请/拒绝邀请(由被邀请人操作)
			If iState=3 Then
				oblog.Execute "Update oblog_teamusers Set state=" & iState & " Where userid=" & oblog.l_uid & " And teamid=" & teamid
				sMsg="已接受该用户的邀请"
			Else
				oblog.execute "Delete From oblog_teamusers Where userid=" & oblog.l_uid & " And teamid=" & teamId
				sMsg="已拒绝该用户的邀请"
			End If
		Case 3
			'取消申请(由申请人操作)
			oblog.execute "Delete From oblog_teamusers Where userid=" & oblog.l_uid & " And teamid=" & teamId
			sMsg="您已取消对该" &oblog.CacheConfig(69)& "的加入申请"
		Case 4
			'接受申请/拒绝申请(由管理员操作)
			Set rs=oblog.execute("select teamid From oblog_team Where teamid=" & teamid & " And managerid=" & oblog.l_uid)
			If Not rs.Eof Then
				If iState=3 Then
					oblog.Execute "Update oblog_teamusers Set state=" & iState & " Where userid=" & userid & " And teamid=" & teamid
					sMsg="已接受该用户的申请"
				Else
					oblog.execute "Delete From oblog_teamusers Where userid=" & userid & " And teamid=" & teamId
					sMsg="已拒绝该用户的申请"
				End If
			Else
				sMsg="您无权对该用户的申请进行操作"
			End If
			Set rs=Nothing
	End select
	oblog.ShowMsg sMsg, "user_team.asp?action=memebers&cmd=" & Request("cmd")
End Sub

Sub SaveAddon()
    Dim rs,sType,sField,sTitle,sDesc,sContent
	sField=Request("itype")
	If  sField="1" Then
		sField="links"
		sTitle=oblog.CacheConfig(69)& "友情连接"
		sDesc="你可以在这里放置" &oblog.CacheConfig(69)& "的与其他站点、博客等的连接"
	Else
		sField="announce"
		sTitle=oblog.CacheConfig(69)& "公告"
		sDesc="你可以在这里放置" &oblog.CacheConfig(69)& "的介绍，或者你愿意放上去的任何信息"
	End If
    sContent = FilterJS(oblog.filt_astr(Request.Form("edit"), 20000))
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select "& sField &" from oblog_team where teamid=" & teamid, conn, 1, 3
    rs(0) =  oblog.filtpath(sContent)
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ShowMsg "修改" & sTitle & "成功", ""
End Sub

Sub ShowAddon(itype)
	Dim rs,sField,sTitle,sDesc
	If  itype="1" Then
		sField="links"
		sTItle=oblog.CacheConfig(69)& "友情连接"
		sDesc="你可以在这里放置" &oblog.CacheConfig(69)& "的与其他站点、博客等的连接"
	Else
		sField="announce"
		sTitle=oblog.CacheConfig(69)& "公告"
		sDesc="你可以在这里放置" &oblog.CacheConfig(69)& "的介绍，或者你愿意放上去的任何信息"
	End If

	Set rs = oblog.execute("select " & sField & " from oblog_team where teamid=" & teamid)
%>
<table class="TeamContent" cellpadding="0">
	<tr>
		<td class="t1">
			<div id="chk_idAll">
				<form name="oblogform" method="post" action="user_team.asp?action=saveaddon&itype=<%=itype%>&teamid=<%=teamid%>" <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
				<table class="TeamInfo" cellpadding="0">
					<tr>
						<td>
							<strong><%=sTitle%></strong><%=sDesc%>
						</td>
					</tr>
					<tr>
						<td>
<span id="loadedit" style="font-size:12px;display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> 正在载入编辑器...</span>
<textarea id="edit" name="edit" style="width:60%;height:220px; display:none"><%=Server.HtmlEncode(OB_IIF(rs(0),""))%></textarea >
<%If C_Editor_Type=2 Then Server.Execute C_Editor & "/edit.asp"%>
						</td>
					</tr>
					<tr>
						<td>
							<input type="submit" name="Submit" id="Submit" value="提交修改" />
						</td>
					</tr>
				</table>
				</form>
<%oblog.MakeEditorText "",1,"535","240"%>
			</div>
			<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
		</td>

<%
teamadmin_top()
Set rs = Nothing
End Sub

sub showconfig()'选择群组模板
%>

<table class="TeamContent" cellpadding="0">
	<tr>
		<td class="t1">
			<div id="chk_idAll" style="overflow-y: hidden; ">
				<iframe id="chgClass"  style="width:100%;height:100%;" src="user_template.asp?action=showconfig&teamid=<%=teamid%>" frameborder="0" scrolling="auto"></iframe>
			</div>
			<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
		</td>
<%
teamadmin_top()
end sub

sub modiconfig()'修改群组主模板
%>

<table class="TeamContent" cellpadding="0">
	<tr>
		<td class="t1">
			<div id="chk_idAll" style="overflow-y: hidden; ">
				<iframe id="chgClass"  style="width:100%;height:100%;" src="user_template.asp?action=modiconfig&teamid=<%=teamid%>" frameborder="0" scrolling="auto"></iframe>
			</div>
			<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
		</td>
<%
teamadmin_top()
end Sub

sub modiviceconfig()'修改群组副模板
%>

<table class="TeamContent" cellpadding="0">
	<tr>
		<td class="t1">
			<div id="chk_idAll" style="overflow-y: hidden; ">
				<iframe id="chgClass"  style="width:100%;height:100%;" src="user_template.asp?action=modiviceconfig&teamid=<%=teamid%>" frameborder="0" scrolling="auto"></iframe>
			</div>
			<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
		</td>
<%
teamadmin_top()
end sub

sub bakskin()'备份群组副模板
%>

<table class="TeamContent" cellpadding="0">
	<tr>
		<td class="t1">
			<div id="chk_idAll" style="overflow-y: hidden; ">
				<iframe id="chgClass"  style="width:100%;height:100%;" src="user_template.asp?action=bakskin&teamid=<%=teamid%>" frameborder="0" scrolling="auto"></iframe>
			</div>
			<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
		</td>
<%
teamadmin_top()
end sub

sub teamadmin_top()
%>
		<td class="t2">
			<ul id="teamadmin_top">
				<li<%If divId=11 Then%> class="Selected"<%End If%>><a href="user_team.asp?action=teamadmin&teamid=<%=teamid%>&div=11" >修改<%=oblog.CacheConfig(69)%>信息</a></li>
				<li<%If divId=12 Then%> class="Selected"<%End If%>><a href="user_team.asp?action=listuser&teamid=<%=teamid%>&div=12" ><%=oblog.CacheConfig(69)%>成员管理</a></li>
				<li<%If divId=13 Then%> class="Selected"<%End If%>><a href="user_team.asp?action=announce&teamid=<%=teamid%>&div=13">设置<%=oblog.CacheConfig(69)%>公告</a></li>
				<li<%If divId=14 Then%> class="Selected"<%End If%>><a href="user_team.asp?action=links&teamid=<%=teamid%>&div=14">设置友情链接</a></li>
				<li<%If divId=17 Then%> class="Selected"<%End If%>><a href="user_team.asp?action=showconfig&teamid=<%=teamid%>&div=17">选择<%=oblog.CacheConfig(69)%>模版</a></li>
				<li<%If divId=15 Then%> class="Selected"<%End If%>><a href="user_team.asp?action=modiconfig&editm=1&teamid=<%=teamid%>&div=15">修改<%=oblog.CacheConfig(69)%>主模版</a></li>
				<li<%If divId=16 Then%> class="Selected"<%End If%>><a href="user_team.asp?action=modiviceconfig&editm=1&teamid=<%=teamid%>&div=16">修改<%=oblog.CacheConfig(69)%>副模版</a></li>
				<li<%If divId=18 Then%> class="Selected"<%End If%>><a href="user_team.asp?action=bakskin&teamid=<%=teamid%>&div=18">备份<%=oblog.CacheConfig(69)%>模版</a></li>
			</ul>
		</td>
	</tr>
</table>
<%end sub

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

sub main()%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr2">
			<th>
				<table id="GroupListTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"><%=oblog.CacheConfig(69)%>最近话题</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<table id="GroupList" class="TableList" cellpadding="0">
<%show_grouplist()%>
					</table>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/18.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<!--<%show_mygroup()%>-->
<%
end  sub

sub show_grouplist()
	dim rs,str,sql
	sql="select top 10 a.topic,a.postid,a.lastupdate,a.teamid,c.t_name,d.username,d.userid,d.nickname from oblog_teampost a,oblog_teamusers b,oblog_team c,oblog_user d where a.iDepth=0 and b.userid="&oblog.l_uid&" and a.teamid=b.teamid and a.teamid=c.teamid and c.istate=3 and (b.state=3 or b.state=5) and a.userid=d.userid order by a.postid desc"
	set rs=oblog.execute(sql)
	while not rs.eof
%>
						<tr>
							<td class="t1">
							</td>
							<td class="t2">
								<span><a href="group.asp?gid=<%=rs("teamid")%>" title="<%=rs("t_name")%>" target="_blank">[<%=rs("t_name")%>]</a></span>&nbsp;-&nbsp;<a href="group.asp?gid=<%=rs("teamid")%>&pid=<%=rs("postid")%>" title="<%=rs("topic")%>
								<%=ob_iif(rs("username"),rs("nickname"))%> 发表于 - <%=FmtMinutes(rs("lastupdate"))%>前" target="_blank"><%=rs("topic")%></a>
							</td>
							<td class="t3">
								<div><a href="go.asp?userid=<%=rs("userid")%>" title="<%=ob_iif(rs("username"),rs("nickname"))%>" target="_blank"><%=ob_iif(rs("username"),rs("nickname"))%></a>&nbsp;发表于&nbsp;-&nbsp;<!--<%=FmtMinutes(rs("lastupdate"))%>前--><span class="time"><%=formatdatetime(rs("lastupdate"),0)%></span></div>
							</td>
						</tr>
<%
		rs.movenext
	wend
	Response.Write(str)
	set rs=nothing
end sub

sub show_mygroup()
	dim rs,str,sql
	set rs=oblog.execute("select a.teamid,a.t_name,a.createtime,a.istate,a.icount0,a.icount1,a.icount2,a.t_ico,b.state From oblog_team  a,oblog_teamusers b Where b.userid=" & oblog.l_uid&" and a.teamid=b.teamid and b.state>2 and a.istate>0 order by b.state desc ")
	while not rs.eof
		str=str&"<ul>"
		str=str&"<li class=""left""><a href=""group.asp?gid="&rs("teamid")&""" title=""点击查看" &oblog.CacheConfig(69)& "："&rs("t_name")&""" target=""_blank""><img src='"&rs("t_ico")&"' /></a></li>"
		str=str&"<li class=""right"">"
		str=str&"	<ol>"
		str=str&"		<li class=""o1""><a href=""group.asp?gid="&rs("teamid")&""" class=""left"" title=""点击查看" &oblog.CacheConfig(69)& "："&rs("t_name")&""" target=""_blank"">"&rs("t_name")&"</a>"
		select Case Cint(OB_IIF(rs("istate"),2))
			Case 1
			str=str&"<font color=""#0000FF"">待审</font>"
			Case 2
			str=str&"<font color=""#ff0000"">锁定</font>"
			Case 3
				if rs("state")>3 then
					str=str&"<a href=""user_team.asp?action=teamadmin&teamid="&rs("teamid")&"&div=13"" class=""right"">管理</a>"
				end if
		end select
		str=str&"</li>"
		str=str&"		<li class=""o2"">创建时间："&rs("createtime")&"</li>"
		str=str&"		<li class=""o3""><span class=""left"">贴数："&rs("icount2")&"篇</span><span class=""right"">成员："&rs("icount0")&"人</span></li>"
		str=str&"	</ol>"
		str=str&"</li>"
		str=str&"</ul>"
		rs.movenext
	wend
	Response.Write(str)
	set rs=nothing
end Sub
'转让群主
sub teammanager()
	dim username,rs,id,teamid,sql,str
	username=oblog.filt_badstr(Trim(Request.Form("t1")))
	teamid=CLng(Request.QueryString("teamid"))
	sql="select userid from oblog_user where username='"&username&"'"
	set rs=oblog.execute(sql)
	If rs.Eof Then
		str="用户名"&username&"不存在"
	Else
		id=CLng(rs(0))

		Set rs = oblog.execute ("select * from oblog_teamusers where teamid= " & teamid & " And userid = " &id)

		If rs.eof Then
			oblog.adderrstr ("此用户非本" &oblog.CacheConfig(69)& "成员！")
	        oblog.showusererr
			Exit Sub
		End if
		Dim trs
		If Not IsObject(conn) Then link_database
		Set trs = Server.CreateObject("adodb.recordset")
		trs.open "select * from oblog_team where teamid=" & teamid & " And managerid=" & oblog.l_uid ,conn,1,3

		If Not trs.eof Then
			trs("managerid")= id
			trs("managername") = username
			trs.update
			oblog.execute ("update oblog_teamusers set state= 5 where teamid=" & teamid & " And userid=" & id)
			str="转让群主成功！"
		Else
			trs.close
			oblog.adderrstr ("无权操作！")
	        oblog.showusererr
			Exit Sub
		End If
		Set trs=Nothing
	End If
	oblog.ShowMsg str,""
end sub
%>
