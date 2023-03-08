<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<!-- #include file="../API/Class_API.asp" -->
<%
If CheckAccess("r_user_all")=False Then Response.Write "无权操作":Response.End
dim rs, sql,rsGroup,sGroups,allGroups,sGroupIds,SqlGroup,SqlGroup2,SqlGroup3
dim UserID,cmd,Keyword,sField,sMail,sMobile,sClass,rsClass
keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
sField=Trim(Request("Field"))
cmd=Trim(Request("cmd"))
Action=Trim(Request("Action"))
UserID=Trim(Request("UserID"))
'ComeUrl=Request.ServerVariables("HTTP_REFERER")

If Session("roleid")<>"" And Session("roleid")<>"0" Then
	Set rs=oblog.Execute("select r_groups From oblog_roles Where roleid=" & Session("roleid"))
	If Not rs.Eof Then sGroupIds=rs(0)
	Set rs=Nothing
End If
If Right(sGroupIds,1)="," Then sGroupIds=Left(sGroupIds,Len(sGroupIds)-1)
If sGroupIds<>"" Then
	If sGroupIds="," Then
		SqlGroup=""
		SqlGroup2=""
		SqlGroup3=""
	Else
		SqlGroup="  Where user_group In (" & sGroupIds & ") "
		SqlGroup2=" And user_group In (" & sGroupIds & ") "
		SqlGroup3=" Where groupid In (" & sGroupIds & ")"
	End If
End If

'整页级别(默认admin)
Set rsClass=oblog.Execute("select id,classname From oblog_userclass  Order By id asc")
Do While Not rsClass.Eof
	sClass=sClass & "<option value="&rsClass(0)&">" & rsClass(1) & "</option>" & vbcrlf
 	rsClass.MoveNext
Loop
Set rsGroup=oblog.Execute("select groupid,g_name From oblog_groups "&SqlGroup3&"  Order By Groupid Desc")
Do While Not rsGroup.Eof
	allGroups=allGroups&rsGroup(0)&"!!??(("&rsGroup(1)&"##))=="
	sGroups=sGroups & "<option value="&rsGroup(0)&">" & rsGroup(1) & "</option>" & vbcrlf
 	rsGroup.MoveNext
Loop
rsGroup.MoveFirst

if cmd="" then
	cmd=0
else
	cmd=CLng(cmd)
end if
G_P_FileName="m_user.asp?cmd=" & cmd
if sField<>"" then
	G_P_FileName=G_P_FileName&"&Field="&sField
end if
if keyword<>"" then
	G_P_FileName=G_P_FileName&"&keyword="&keyword
end if
if Request("page")<>"" then
    G_P_This=cint(Request("page"))
else
	G_P_This=1
end if
If cmd = 101 Then
	G_P_FileName = G_P_FileName & "&groupid="&clng(Request("groupid"))
End if
If cmd = 109 Then
	G_P_FileName=G_P_FileName&"&ClassID="&CLng(Request("classid"))
End If
%>
<script language=javascript>
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
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<%If action<>"Update" Then%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">注 册 用 户 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_user.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查询：</strong></td>
      <td width="687" height="30">
      	<select size=1 name="cmd">
          <option value=>请选择查询条件</option>
		  <option value="1">最后注册的500个用户</option>
		  <option value="2">最早注册的500个用户</option>
          <option value="3">文章最多的100个用户</option>
          <option value="4">文章最少的100个用户</option>
		  <option value="5">推荐博客</option>
		  <option value="6">所有待审核用户</option>
<!--           <option value="7">等待管理员认证的用户</option> -->
          <option value="8">所有被锁住的用户</option>
		  <option value="10">所有被前台屏蔽的用户</option>
        </select>
        <input type="submit" value=" 查 询 ">
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="m_user.asp">用户管理首页</a>&nbsp;|&nbsp;<a href="m_user.asp?cmd=6">待审核用户</a>|&nbsp;<a href="m_user.asp?cmd=9"><font color=red>敏感用户</font></a>|&nbsp;<a href="m_user.asp?cmd=10">被前台屏蔽的用户</a>|&nbsp;<a href="../reg.asp" target="_blank">添加新用户</a></td>
    </tr>
  </form>
  <form name="form2" action="m_user.asp?cmd=101" method="post">
    <tr class="tdbg">
      <td width="100" height="30"><strong>按用户组查询：</strong></td>
      <td width="687" height="30">
      	<select size=1 name="groupid">
      	  <option value="0">------尚未分组------</option>
          <%=sGroups%>
        </select>
        <input type="submit" value=" 查 询 "></td>
    </tr>
  </form><%If oblog.filt_badstr(session("adminname"))<>"" Then %>
   <form name="form2" action="m_user.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>按用户分类查询：</strong></td>
      <td width="687" height="30">
      	<select size=1 name="classid">
      	  <option value="0">------全部分类------</option>
          <%=sClass%>
        </select>
		 <input name="cmd" type="hidden" id="cmd" value="109">
        <input type="submit" value=" 查 询 "></td>
    </tr>
  </form><%End If %>
  <form name="form3" method="post" action="m_user.asp">
  <tr class="tdbg">
    <td width="120"><strong>用户高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">
		  <option value="UserName" selected>用户名</option>
	      <option value="UserID">用户ID</option>
		  <option value="nickname">用户昵称</option>
		  <option value="blogname">blog名称</option>
		  <option value="email">注册用Email</option>
		  <option value="regip">注册用ip</option>
		  <option value="regdate">注册时间(格式YYYYMMDD,如20060601)</option>
		  <option value="birthday">生日(格式YYYYMMDD,如20060601)</option>
		  <option value="regcity">所在省市(请手工填写,省市之间用,隔开,如山东,威海)</option>
	      <option value="loginip" >最后登录ip</option>
		  <option value="lastlogintime" >多少天内未登录</option>
		  <option value="logcount">文章数小于</option>
		  <option value="logintimes">登录次数小于</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="cmd" type="hidden" id="cmd" value="102">
	  若为空，则查询所有用户</td>
  </tr>
</form>
<tr><td colspan=2><b>内容管理员仅可管理自己被许可管理的用户等级。但敏感用户、锁定用户、待审用户不受限<br/>
用户积分重计计算方法：注册初始分+日志分+精华分+回复分+留言分+群组帖子分+创建圈子奖励分-创建圈子消耗<br/>
重计可能不准确，只能根据现有数据统计，不能统计因为删除等操作扣除的积分等过往情况	</b>
</td></tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%End If%>
<%
Select Case LCase(action)
 Case   "add"
    call adduser
 Case   "saveadd"
    call saveadd
 Case   "modify"
    call modify
 Case   "savemodify"
    call savemodify
 Case   "del"
    call deluser
 Case   "lock"
    call lockuser
 Case   "unlock"
    call unlockuser
 Case   "block"
    call blockuser
 Case   "unblock"
    call unblockuser
 Case   "move"
    call moveuser
 Case   "update"
    call updateuser
 Case   "doupdate"
    call doupdate
 Case   "doupdatelog"
    call doupdatelog
 Case   "gouser1"
    call gouser1
 Case   "gouser2"
    call gouser2
 Case   "pass"
    call passit(7)
 Case   "unpass"
    call passit(6)
 Case  "rescore"
	call rescore
Case  else
    call main
End Select
If FoundErr = True Then
    Call WriteErrMsg
End If

Sub main()
    Dim QryFields
	Dim sDate
    sGuide=""
    QryFields=" top 500 userid,username,user_icon1,regip,adddate,lockuser,user_level,user_group,lastloginip,lastlogintime,logintimes,istrouble,emailvalid,log_count,is_log_default_hidden "
    select Case cmd
        Case 1
            sql = "select " & QryFields &" from oblog_user " & SqlGroup &" order by UserID desc"
            sGuide = sGuide & "最后注册的500个用户"
        Case 2
            sql = "select  " & QryFields &"  from oblog_user " & SqlGroup &" order by UserID"
            sGuide = sGuide & "最早注册的500个用户"
        Case 3
            sql = "select  " & QryFields &"  from  oblog_user " & SqlGroup &" order by log_count Desc"
            sGuide = sGuide & "发表日志最多的500个用户"
        Case 4
            sql = "select   " & QryFields &"  from  oblog_user " & SqlGroup &" order by log_count"
            sGuide = sGuide & "发表日志最少的500个用户"
        Case 5
            sql = "select  " & QryFields &"  from  oblog_user where user_isbest=1 " & SqlGroup2 &" order by userid desc"
            sGuide = sGuide & "推荐博客"
        Case 6
            sql = "select   " & QryFields &"  from  oblog_user where User_Level=6 order by userid desc"
            sGuide = sGuide & "等待管理审核的用户"
        Case 8
            sql = "select  " & QryFields &"  from oblog_user where  LockUser =1 order by userID  desc"
            sGuide = sGuide & "被锁住的用户"
        Case 9
            sql = "select   " & QryFields &"  from oblog_user where  istrouble >0 order by userID  desc"
            sGuide = sGuide & "<font color=red>敏感用户(任何发布过敏感/反动关键字的用户将被归入此类)</font>"
		Case 10
            sql = "select  " & QryFields &"  from oblog_user where  is_log_default_hidden =1 order by userID  desc"
            sGuide = sGuide & "被系统前台屏蔽锁的用户"

		Case 109
			If oblog.filt_badstr(session("adminname"))<>"" Then
			sql = "select  " & QryFields &"  from oblog_user where user_classid="&clng(Request("classid"))
			sGuide = sGuide & "按会员分类查询"
			Else
			sql = "select  " & QryFields &"  from oblog_user where 1=2"
			sGuide = sGuide & "内容管理员无权按会员分类查询"
			End If
        Case 101
            If Request("groupid") = 0 Then
                sql = "select   " & QryFields &"  from oblog_user where  user_group is null " & SqlGroup2
            Else
                sql = "select   " & QryFields &"  from oblog_user where  user_group=" & clng(Request("groupid"))& SqlGroup2
            End If
            sGuide = sGuide & "按会员组查询"

        Case 102
            If Keyword = "" Then
                sql = "select   " & QryFields &"  from oblog_user " & SqlGroup &" order by userID desc"
				'sGuide = sGuide & "所有用户"
				sGuide = sGuide & "最新注册的500个用户"
            Else
                select Case LCase(sField)
                Case "userid"
                    If IsNumeric(Keyword) = False Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<br><li>用户ID必须是整数！</li>"
                    Else
                        sql = "select  " & QryFields &"  from oblog_user where userID =" & CLng(Keyword)  & SqlGroup2
                        sGuide = sGuide & "用户ID等于<font color=red> " & CLng(Keyword) & " </font>的用户"
                    End If
                Case "username"
                    'If is_sqldata = 1 Then
                        sql = "select  " & QryFields &"  from oblog_user where username like '%" & Keyword & "%' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "用户名中含有“ <font color=red>" & Keyword & "</font> ”的用户"
                    'Else
                    '    sql = "select  " & QryFields &"  from oblog_user where username= '" & Keyword & "' " & SqlGroup2 &" order by userID  desc"
                    '    sGuide = sGuide & "用户名等于“ <font color=red>" & Keyword & "</font> ”的用户"
                    'End If

                Case "nickname"
                    If is_sqldata = 1 Then
                        sql = "select  " & QryFields &"  from oblog_user where nickname like '%" & Keyword & "%' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "用户昵称中含有“ <font color=red>" & Keyword & "</font> ”的用户"
                    Else
                        sql = "select  " & QryFields &"  from oblog_user where nickname='" & Keyword & "' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "用户昵称等于“ <font color=red>" & Keyword & "</font> ”的用户"
                    End If
                Case "regip"
                    If is_sqldata = 1 Then
                        sql = "select  " & QryFields &"  from oblog_user where regip like '%" & Keyword & "%' " & SqlGroup2 &"  order by userID  desc"
                        sGuide = sGuide & "注册ip中含有“ <font color=red>" & Keyword & "</font> ”的用户"
                    Else
                        sql = "select  " & QryFields &"  from oblog_user where regip='" & Keyword & "' " & SqlGroup2&" order by userID  desc"
                        sGuide = sGuide & "注册ip等于“ <font color=red>" & Keyword & "</font> ”的用户"
                    End If
                Case "loginip"
                    If is_sqldata = 1 Then
                        sql = "select  " & QryFields &"  from oblog_user where lastloginip like '%" & Keyword & "%' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "最后登录ip中含有“ <font color=red>" & Keyword & "</font> ”的用户"
                    Else
                        sql = "select  " & QryFields &"  from oblog_user where lastloginip='" & Keyword & "' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "最后登录ip等于“ <font color=red>" & Keyword & "</font> ”的用户"
                    End If
                Case "blogname"
                    If is_sqldata = 1 Then
                        sql = "select  " & QryFields &"  from oblog_user where blogname like '%" & Keyword & "%' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "blog名中含有“ <font color=red>" & Keyword & "</font> ”的用户"
                    Else
                        sql = "select  " & QryFields &"  from oblog_user where blogname='" & Keyword & "' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "blog名等于“ <font color=red>" & Keyword & "</font> ”的用户"
                    End If
                Case "logcount"
                    sql = "select   " & QryFields &"  from oblog_user where log_count < " & CLng(Keyword) & " " & SqlGroup2 &" order by userID  desc"
                    sGuide = sGuide & "文章数小于“ <font color=red>" & Keyword & "</font> ”的用户"
                Case "logintimes"
                    sql = "select   " & QryFields &"  from oblog_user where logintimes < " & CLng(Keyword) & " " & SqlGroup2 &"order by userID  desc"
                    sGuide = sGuide & "登录次数小于“ <font color=red>" & Keyword & "</font> ”的用户"
                Case "lastlogintime"
					sql = "select   " & QryFields &"  from oblog_user where "
					If Is_Sqldata = 0 Then
						sql = sql & " datediff("&G_Sql_d&",lastlogintime,"&G_Sql_Now&")>" & Int(Keyword) & SqlGroup2 &" order by userID  desc"
					Else
						sDate = DateAdd ("d",-1*Abs(Keyword),Date())
						sDate = GetDateCode(sDate,0)
						sql = sql & " lastlogintime < '"&sDate& SqlGroup2 &"' ORDER BY userid DESC"
					End if
                    sGuide = sGuide & "“ <font color=red>" & Keyword & "</font> ”天内未登录的用户"
                    'New
                Case "email"
                    If is_sqldata = 1 Then
                        sql = "select  " & QryFields &"  from oblog_user where useremail like '%" & Keyword & "%' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "注册Email名中含有“ <font color=red>" & Keyword & "</font> ”的用户"
                    Else
                        sql = "select  " & QryFields &"  from oblog_user where useremail='" & Keyword & "' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "注册Email等于“ <font color=red>" & Keyword & "</font> ”的用户"
                    End If
                Case "regdate"
					sDate = keyword
                	Keyword=DeDateCode(keyword)
                	If Keyword<>"" Then
                        sql = "select  " & QryFields &"  from oblog_user where "
						If Is_Sqldata = 0 Then
							sql = sql & " datediff(" & G_Sql_d & ",adddate,'" & keyword & "')=0 " & SqlGroup2 &" Order By userid Desc "
						Else
							sql = sql & " adddate = '"&sDate&"' " & SqlGroup2&"  ORDER BY userid DESC "
						End if
                        sGuide = sGuide & "注册日期为“ <font color=red>" & Keyword & "</font> ”的用户"
              		End If
                Case "birthday"
					sDate = keyword
					Keyword=DeDateCode(keyword)
                	If Keyword<>"" Then
                        sql = "select  " & QryFields &"  from oblog_user where "
						If Is_Sqldata = 0 Then
							sql = sql & " datediff(" & G_Sql_d & ",birthday,'" & keyword & "')=0 " & SqlGroup2 &"  Order By userid Desc "
						Else
							sql = sql & " birthday = '"&sDate&"' " & SqlGroup2 &" ORDER BY userid DESC "
						End if
                        sGuide = sGuide & "生日为“ <font color=red>" & Keyword & "</font> ”的用户"
              		End If
                Case "regcity"
                	Dim aCity
					keyword=Replace (keyword,"，",",")
                	If InStr(keyword,",")>0 Then
                		aCity=Split(keyword,",")
                		sql = "select  " & QryFields &"  from oblog_user where province='" & aCity(0)&"' And city='"& aCity(1)&"' " & SqlGroup2 &" order by userID  desc"
                        sGuide = sGuide & "所在地为“ <font color=red>" & Keyword & "</font> ”的用户"
                    End If
                End select
            End If
        Case Else
            sql = "select   " & QryFields &"  from oblog_user " & SqlGroup &" order by UserID desc"
            sGuide = sGuide & "最后注册的500个用户"
    End select

    If FoundErr = True Then Exit Sub
	If sql = "" Then Oblog.ShowMsg "格式不正确，请重试",""
'	OB_DEBUG sql,1
    If Not IsObject(Conn) Then link_database
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.open sql, Conn, 1, 1
    If rs.EOF And rs.bof Then
        sGuide = sGuide & "(<font color=red>0</font>)"
        %>
        <div id="main_body">
		<ul class="main_top">
			<li class="main_top_left left"><%=sGuide%></li>
			<li class="main_top_right right"> </li>
		</ul>
		</div>
        <%
    Else
        G_P_AllRecords = rs.recordcount
        sGuide = sGuide & "(<font color=red>" & G_P_AllRecords & "</font>)"
        If G_P_This < 1 Then
            G_P_This = 1
        End If
        If (G_P_This - 1) * G_P_PerMax > G_P_AllRecords Then
            If (G_P_AllRecords Mod G_P_PerMax) = 0 Then
                G_P_This = G_P_AllRecords \ G_P_PerMax
            Else
                G_P_This = G_P_AllRecords \ G_P_PerMax + 1
            End If

        End If
        If G_P_This = 1 Then
            showContent
            Response.Write oblog.showpage(True, True, "个用户")
        Else
            If (G_P_This - 1) * G_P_PerMax < G_P_AllRecords Then
                rs.Move (G_P_This - 1) * G_P_PerMax
                Dim bookmark
                bookmark = rs.bookmark
                showContent
                Response.Write oblog.showpage(True, True, "个用户")
            Else
                G_P_This = 1
                showContent
                Response.Write oblog.showpage(True, True, "个用户")
            End If
        End If
    End If
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
  <form name="myform" method="Post" action="m_user.asp" onsubmit="return confirm('确定要执行选定的操作吗？');">
<style type="text/css">
<!--
.border tr td {padding:3px 0!important;}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="title">
    <td align="center" width="28"><strong>选中</strong></td>
    <td align="center" width="44"><strong>ID</strong></td>
    <td align="center" width="58"><strong>用户头像</strong></td>
    <td align="center"><strong>用户名 用户组</strong></td>
<!--     <td align="center" width="58"><strong>邮箱认证</strong></td> -->
    <td align="center" width="100"><strong>注册时间 注册IP</strong></td>
    <td align="center" width="100"><strong>登录时间 登录IP</strong></td>
    <td align="center" width="58"><strong>登录数</strong></td>
    <td align="center" width="58"><strong>日志数</strong></td>
    <td align="center" width="70"><strong>审核 锁定</strong></td>
    <td align="center" width="100"><strong>操作</strong></td>
  </tr>
          <%do while not rs.EOF %>
  <tr class="title">
    <td align="center"><input name='UserID' type='checkbox' onclick="unselectall()" id="UserID" value='<%=cstr(rs("userID"))%>'></td>
    <td align="center" style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("userID")%></td>
    <td align="center"><a href="../blog.asp?name=<%=rs("userName")%>" target="_blank" title="点击访问该博客"><img src="<%=ProIco(rs("user_icon1"),1)%>" align="absmiddle" style="width:48px;height:48px;border:0;"></a></td>
    <td style="padding:0 0 0 3px!important;">
	<a href="../blog.asp?name=<%=rs("userName")%>" target="_blank" style="font-weight:400;"  title="点击访问该博客"><%=rs("userName")%></a>

	<span style="display:block;color:#217DBD;">
	<%=GetsubName(rs("user_group"),allGroups)%>
	</span>
	</td>
<!--     <td align="center">
            <%
		    select Case OB_IIF(rs("emailValid"),"0")
		    	Case "0"
		    		Response.Write "<font color=#FF6600>未认证</font>"
		    	Case "1"
		    		Response.Write "<font color=#009900>已认证</font>"
		    End select
    		%>
	</td> -->
    <td style="color:#999;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:10px;padding:0 0 0 8px!important;">
	<%=OB_IIF(rs("adddate"),"&nbsp;") %>
	<br />
	<%=OB_IIF(rs("regip"),"&nbsp;") %>
	</td>
    <td style="color:#666;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:10px;padding:0 0 0 8px!important;">
	<%=OB_IIF(rs("LastLoginTime"),"&nbsp;") %>
	<br />
	<%=OB_IIF(rs("lastloginip"),"-") %>
	</td>
    <td align="center"><strong><%=OB_IIF(rs("LoginTimes"),"0") %></strong></td>
    <td align="center"><strong><%=OB_IIF(rs("log_count"),"0") %></strong></td>
    <td align="center">
	<%
		    select Case OB_IIF(rs("user_level"),"6")
		    	Case "6"
		    		Response.Write "<span style=""color:#f00;font-weight:600;"">待审</span>"
		    	Case "7"
		    		Response.Write "<span style=""color:#090;font-weight:600;"">已审</span>"
		    End select
    %>&nbsp;
      <%
      If rs("user_level") = 6 Then
      	Response.Write "<span><a href='m_user.asp?Action=pass&UserID=" & rs("userID") & "' title=""通过审核"">审核</a></span>  "
      Else
      	Response.Write "<span><a href='m_user.asp?Action=unpass&UserID=" & rs("userID") & "' title=""取消审核"">取消</a></span>  "
      End If
        %>
	<br />
	  <%
      If rs("LockUser") = 1 Then
        Response.Write "<span style=""color:#f00;font-weight:600;"">已锁</span>  "
      Else
        Response.Write "<span style=""color:#090;font-weight:600;"">正常</span>  "
      End If
      %>
		<%
        If rs("LockUser") = 0 Then
            Response.Write "<a href='m_user.asp?Action=Lock&UserID=" & rs("userID") & "'>锁定</a>"
        Else
            Response.Write "<a href='m_user.asp?Action=UnLock&UserID=" & rs("userID") & "'>解锁</a>"
        End If
        %>
	</td>
    <td align="left">
		<%
        Response.Write "&nbsp;<a href='m_user.asp?Action=Modify&UserID=" & rs("userID") & "'>修改</a>&nbsp;"
        If CheckAccess("r_user_admin") Then
	        Response.Write "<a href='m_user.asp?Action=gouser2&username=" & rs("username") & "' target='blank'>进后台</a>&nbsp;"
	    End If
        Response.Write "<a href='m_user.asp?Action=Del&UserID=" & rs("userID") & "' onClick='return confirm(""确定要删除此用户吗？"");'>删除</a>&nbsp;"
        Response.Write "<br/>&nbsp;"
        Response.Write "<a href='m_user.asp?Action=rescore&UserID=" & rs("userID") & "'>修复用户</a>&nbsp;"
		If CheckAccess("r_user_admin") Then
			If rs("is_log_default_hidden") = 1 Then
				Response.Write "<a href=""m_user.asp?Action=Unblock&UserID=" & rs("userID") & """ style=""color:red;"">取消屏蔽</a>&nbsp;"
			Else
				Response.Write "<a href=""m_user.asp?Action=block&UserID=" & rs("userID") & """ style=""color:green;"">前台屏蔽</a>&nbsp;"
			End If
		End If
        %>
	</td>
  </tr>
<%
    i = i + 1
    If i >= G_P_PerMax Then Exit Do
    rs.movenext
Loop
%>
        </table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有用户</td>
            <td> <strong>操作：</strong>
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Group.disabled=true">
              删除&nbsp;&nbsp;&nbsp;&nbsp;
              <input name="Action" type="radio" value="Move" onClick="document.myform.User_Group.disabled=false">移动到
              <select name="User_Group" id="User_Group" disabled>
                <%=sGroups%>
              </select>
              &nbsp;&nbsp;
              <input name="Action" type="radio" value="unpass" checked onClick="document.myform.User_Group.disabled=true">
              待审&nbsp;&nbsp;&nbsp;&nbsp;
              <input name="Action" type="radio" value="pass" checked onClick="document.myform.User_Group.disabled=true">
              审核&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="submit" name="Submit" value=" 执 行 "> </td>
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


Sub Modify()
    Dim userid
    Dim rsUser, sqlUser
	Dim sql
    userid = Trim(Request("UserID"))
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>参数不足！</li>"
        Exit Sub
    Else
        userid = CLng(userid)
    End If
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
	sql = "userid,username,user_domain,user_domainroot,blogname"&str_domain&",user_classid,Question,Sex,userEmail,qq,Msn,User_Group,scores,user_upfiles_size,user_isbest,user_dir,LockUser,User_Level"
    sqlUser = "select "&sql&" from oblog_user where userID=" & userid
    If Not IsObject(Conn) Then link_database
    rsUser.open sqlUser, Conn, 1, 3
    If rsUser.bof And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>找不到指定的用户！</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">修改注册用户信息</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<FORM name="Form1" action="m_user.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class="tdbg" >
      <TD width="40%">用户名：</TD>
      <TD width="60%"><%=rsUser("userName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td>用户域名：</td>
      <td><input name="user_domain" type="text" value="<%=oblog.filt_html(rsuser("user_domain"))%>" size=10 maxlength=20 /> <select name="user_domainroot" ><%=oblog.type_domainroot(rsuser("user_domainroot"),0)%></select></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td>blog名：</td>
      <td><input name=blogname   type=text id="blogname" value="<%=rsuser("blogname")%>" size=30 maxlength=20></td>
    </tr>
    <%if true_domain=1 then%>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>用户绑定的顶级域名：</td>
      <td><input name=custom_domain   type=text id="custom_domain" value="<%=rsuser("custom_domain")%>" size=30 maxlength=20></td>
    </tr>
    <%end if%>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td>blog类别：</td>
      <td><select name="usertype" id="usertype">
          <%if rsUser("user_classid")<>"" then
      Response.Write (oblog.show_class("user", rsUser("user_classid"), 0))
      Else
      Response.Write (oblog.show_class("user", 0, 0))
      End If
      %>
        </select></td>
    </tr>
    <TR class="tdbg" >
      <TD width="40%">密码(至少6位)：<BR>
        请输入密码，区分大小写。 请不要使用任何类似 '*'、' ' 或 HTML 字符 </TD>
      <TD width="60%"> <INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">如果不想修改，请留空(整合用户请到论坛修改)</font> </TD>
    </TR>
    <TR class="tdbg" >
      <TD>确认密码(至少6位)：<br>
        请再输一遍确认</TD>
      <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=16> <font color="#FF0000">如果不想修改，请留空(整合用户请到论坛修改)</font> </TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">密码问题：<br>
        忘记密码的提示问题</TD>
      <TD width="60%"> <INPUT name="Question"   type=text value="<%=rsUser("Question")%>" size=30>(整合用户请到论坛修改)
      </TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">问题答案：<BR>
        忘记密码的提示问题答案，用于取回密码</TD>
      <TD width="60%"> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">如果不想修改，请留空(整合用户请到论坛修改)</font></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">性别：</TD>
      <TD width="60%"> <INPUT type=radio value="1" name=sex <%if rsUser("Sex")=1 then Response.write "CHECKED"%>>
        男 &nbsp;&nbsp; <INPUT type=radio value="0" name=sex <%if rsUser("Sex")=0 then Response.write "CHECKED"%>>
        女</TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">Email地址：</TD>
      <TD width="60%"> <INPUT name=Email value="<%=rsUser("userEmail")%>" size=30   maxLength=50>
        <a href="mailto:<%=rsUser("userEmail")%>">给此用户发一封电子邮件</a>
      </TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">OICQ号码：</TD>
      <TD width="60%"> <INPUT name=OICQ value="<%=rsUser("qq")%>" size=30 maxLength=20></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">MSN：</TD>
      <TD width="60%"> <INPUT name=msn value="<%=rsUser("Msn")%>" size=30 maxLength=50></TD>
    </TR>
    <%If CheckAccess("r_user_group") Then%>
    <TR class="tdbg" >
      <TD width="40%"><font color=red><b>用户组</b></font>：</TD>
      <TD width="60%">
      	<select name="groupid" id="groupid">
          <%
          Dim rsGroup,userGroup
          Set rsGroup=oblog.Execute("select groupid,g_name,g_level From oblog_groups Order By g_level")
          userGroup=Int(OB_IIF(rsUser("User_Group"),0))
          If userGroup=0 Then%>
          		<option value="0" selected>----尚未分类----</option>
          <%End If
          Do While Not rsGroup.Eof%>
          	<option value="<%=rsGroup(0)%>" <%If rsGroup(0)=UserGroup Then%>selected<%End if%>><%=rsGroup(2)%>-<%=rsGroup(1)%></option>
          		<%
          	rsGroup.Movenext
        	Loop
        	Set rsGroup=Nothing
          %>
        </select>(如果是将用户降级,则请同步修改(减少)积分)</TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%"><font color=red><b>用户积分</b></font>：</TD>
      <TD width="60%"> <INPUT name=scores value="<%=rsUser("scores")%>" size=30 maxLength=10></TD>
    </TR>
    <%End If%>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td>已上传字节(字节)：</td>
      <td><input name=upfiles_size   type=text id="upfiles_size" value="<%=rsuser("user_upfiles_size")%>" size=30 maxlength=20></td>
    </tr>
    <TR class="tdbg" >
      <TD>是否为推荐博客：</TD>
      <TD><input type="radio" name="isbest" value=1 <%if rsUser("user_isbest")=1 then Response.write "checked"%>>
        是 &nbsp;&nbsp; <input type="radio" name="isbest" value=0 <%if rsUser("user_isbest")<>1 then Response.write "checked"%>>
        否</TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">用户目录：</TD>
      <TD width="60%"> <INPUT name=user_dir value="<%=rsUser("user_dir")%>" size=30 maxLength=50>
        如无必要请不要修改，否则将造成用户目录混乱</TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">用户状态：</TD>
      <TD width="60%"><input type="radio" name="LockUser" value=0 <%if rsUser("LockUser")=0 then Response.write "checked"%>>
        正常&nbsp;&nbsp; <input type="radio" name="LockUser" value=1 <%if rsUser("LockUser")=1 then Response.write "checked"%>>
        锁定</TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">审核状态：</TD>
      <TD width="60%"><input type="radio" name="User_Level" value=6 <%if rsUser("User_Level")=6 then Response.write "checked"%>>
        未审核&nbsp;&nbsp; <input type="radio" name="User_Level" value=7 <%if rsUser("User_Level")=7 then Response.write "checked"%>>
        已审核</TD>
    </TR>
    <%If oblog.cacheConfig(51)="1"  Then

    	If Not IsNull(rsuser("postmail")) Then sMail=rsuser("postmail")
    	If Not IsNull(rsuser("postmobile")) Then sMobile=rsuser("postmobile")
    	%>
    <TR class="tdbg" >
      <TD width="40%">发布用邮箱地址：</TD>
      <TD width="60%"> <INPUT   type=text maxLength=100 size=30 name=postmail value="<%=sMail%>"> <font color="#FF0000"></font> </TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">发布用手机号码：         </TD>
      <TD width="60%"> <INPUT   type=text maxLength=11 size=30 name=postmobile value="<%=sMobile%>"> <font color="#FF0000">目前只支持中国移动GSM号码</font> </TD>
    </TR>
    <%End If%>
    <TR class="tdbg" >
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="保存修改结果"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("userID")%>"></TD>
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
    rsUser.Close
    Set rsUser = Nothing
End Sub


Sub UpdateUser()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">更 新 页 面</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<FORM name="Form1" action="m_user.asp?action=DoUpdate" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title">
      <td height="22" colspan="2" class="topbg"><strong>更新用户静态页面</font></strong></td>
  </tr>
  <tr class="tdbg">
      <td colspan="2"><p>说明：<br>
          1、本操作将重新生成用户静态页面。<br>
          2、本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。<br>
      3 、本操作根据用户ｉｄ更新。 </p>
      </td>
  </tr>
  <tr class="tdbg">
    <td height="25">开始用户ID：</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      用户ID，可以填写您想从哪一个ID号开始进行更新</td>
  </tr>
  <tr class="tdbg">
    <td height="25">结束用户ID：</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
      将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
  </tr>
     <tr class="tdbg">
    <td height="25">&nbsp;</td>
    <td height="25"><input type="checkbox" name="t0" value="1" /><font color="red"><b>同时满足以下条件</b></font></td>
  </tr>
  <tr class="tdbg">
    <td height="25">登录次数大于：</td>
    <td height="25"><input name="Logintimes" type="text" id="Logintimes" value="0" size="10" maxlength="10">（不包括指定数值）
</td>
  </tr>
  <tr class="tdbg">
    <td height="25">日志数：</td>
    <td height="25"><input name="B_Logs" type="text" id="B_Logs" value="0" size="10" maxlength="10">&nbsp;至&nbsp;<input name="E_Logs" type="text" id="E_Logs" value="1000" size="10" maxlength="10">
    &nbsp; </td>
  </tr>
  <tr class="tdbg">
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="生成静态页面"></td>
  </tr>
</table>
</form>
<FORM name="Form1" action="m_user.asp?action=DoUpdatelog" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title">
      <td height="22" colspan="2" class="topbg"><strong>更新日志静态页面</font></strong></td>
  </tr>
  <tr class="tdbg">
      <td colspan="2"><p>说明：<br>
          1、本操作将重新生成用户静态页面。<br>
          2、本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。<br>
      3、本操作根据日志ｉｄ更新。</p>
      </td>
  </tr>
  <tr class="tdbg">
    <td height="25">开始日志ID：</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      日志ID，可以填写您想从哪一个ID号开始进行更新</td>
  </tr>
  <tr class="tdbg">
    <td height="25">结束日志ID：</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
      将更新开始到结束ID之间的日志页面，之间的数值最好不要选择过大</td>
  </tr>
  <tr class="tdbg">
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="生成日志静态页面"></td>
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

Sub gouser1()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">登录到用户管理后台</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<FORM name="Form1" action="m_user.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="tdbg">
      <td colspan="2"><p>说明：<br>
          本操作供管理员登录到用户的管理界面进行管理。<br>
          当用户操作出现障碍时，可进入该用户后台，协助用户进行操作。<br>
        </p>
      </td>
  </tr>
  <tr class="tdbg">
    <td height="25">用户账号：</td>
    <td height="25"><input name="username" type="text" id="username" value="" size="30" maxlength="50"></td>
  <tr class="tdbg">
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value=" 提交 "></td>
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
%>
</body>
</html>
<%

Sub gouser2()
    Dim rs, username
    If CheckAccess("r_user_admin")=false Then
    	 Response.Write("您没有进入用户后台的权限"):Response.End()
	End If
    username = oblog.filt_badstr(Trim(Request("username")))
    if username="" then Response.Write("用户名不能为空"):Response.End()
    set rs=Server.CreateObject("adodb.recordset")
	rs.open("select username,TruePassWord,user_group from oblog_user where username='"&username&"'"),conn,1,3
    If Not rs.EOF Then
		If Not CheckGoUser(rs(2)) Then
			Response.Write "无权限"
			Exit Sub
		End If
		If IsNull(rs(1)) Then
			rs(1) = RndPassword(16)
			rs.update
		End if
        oblog.SaveCookie rs(0), rs(1), 0
        Set rs = Nothing
		WriteSysLog "进行了进入用户后台操作，目标用户："&username&"",oblog.NowUrl&"?"&Request.QueryString
        Response.Redirect ("../user_index.asp")
    Else
        Set rs = Nothing
        Response.Write("无此用户"):Response.End()
    End If
End Sub
Sub SaveModify()
	If Request.QueryString <>"" Then Exit Sub
    Dim userid, Password, PwdConfirm, Question, Answer, Sex, Email, Homepage, OICQ, MSN, User_Level, LockUser, isbest
    Dim rsUser, sqlUser,Scores,user_Group
    Dim blogname, usertype, user_upfiles_max, upfiles_size, user_domain, user_domainroot
    Action = Trim(Request("Action"))
    userid = Trim(Request("UserID"))
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>参数不足！</li>"
        Exit Sub
    Else
        userid = CLng(userid)
    End If
    Password = Trim(Request("Password"))
    PwdConfirm = Trim(Request("PwdConfirm"))
    Question = Trim(Request("Question"))
    Answer = Trim(Request("Answer"))
    Sex = Trim(Request("Sex"))
    Email = Trim(Request("Email"))
    Homepage = Trim(Request("Homepage"))
    OICQ = Trim(Request("OICQ"))
    MSN = Trim(Request("MSN"))
    User_Level = Trim(Request("User_Level"))
    isbest = Trim(Request("isbest"))
    LockUser = Trim(Request("LockUser"))
    blogname = Trim(Request("blogname"))
    usertype = Trim(Request("usertype"))
    user_upfiles_max = Trim(Request("user_upfiles_max"))
    upfiles_size = Trim(Request("upfiles_size"))
    user_domain = Trim(Request("user_domain"))
    user_domainroot = Trim(Request("user_domainroot"))
	user_group= Request("groupid")
	scores= Request("scores")
    If Password <> "" Then
        If oblog.strLength(Password) > 12 Or oblog.strLength(Password) < 6 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>密码不能大于12小于6，如果你不想修改密码，请保持为空。</li>"
        End If
        If InStr(Password, "=") > 0 Or InStr(Password, "%") > 0 Or InStr(Password, Chr(32)) > 0 Or InStr(Password, "?") > 0 Or InStr(Password, "&") > 0 Or InStr(Password, ";") > 0 Or InStr(Password, ",") > 0 Or InStr(Password, "'") > 0 Or InStr(Password, ",") > 0 Or InStr(Password, Chr(34)) > 0 Or InStr(Password, Chr(9)) > 0 Or InStr(Password, "") > 0 Or InStr(Password, "$") > 0 Then
            ErrMsg = ErrMsg + "<br><li>密码中含有非法字符，如果你不想修改密码，请保持为空。</li>"
            FoundErr = True
        End If
    End If
    If Password <> PwdConfirm Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>密码和确认密码不一致</li>"
    End If
    If Question = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>密码提示问题不能为空</li>"
    End If

    If Sex = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>性别不能为空</li>"
    Else
        Sex = CInt(Sex)
        If Sex <> 0 And Sex <> 1 Then
            Sex = 1
        End If
    End If

        If Email = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>Email不能为空</li>"
        Else
            If oblog.IsValidEmail(Email) = False Then
                ErrMsg = ErrMsg & "<br><li>您的Email有错误</li>"
                FoundErr = True
            End If
        End If

    If OICQ <> "" Then
        If Not IsNumeric(OICQ) Or Len(CStr(OICQ)) > 10 Then
            ErrMsg = ErrMsg & "<br><li>OICQ号码只能是4-10位数字，您可以选择不输入。</li>"
            FoundErr = True
        End If
    End If
    If MSN <> "" Then
        If oblog.IsValidEmail(MSN) = False Then
            ErrMsg = ErrMsg & "<br><li>你的MSN有误。</li>"
            FoundErr = True
        End If
    End If
	If CheckAccess("r_user_group") Then
	    If User_Group= 0 Then
	        FoundErr = True
	        ErrMsg = ErrMsg & "<br><li>请指定用户级别！</li>"
	    Else
	        User_Group = CLng(User_Group)
	    End If

	    If Not IsNumeric(scores) Then
	        FoundErr = True
	        ErrMsg = ErrMsg & "<br><li>请正确填写用户积分！</li>"
	    Else
	        scores = Int(scores)
	    End If
	End If
    If oblog.cacheConfig(4) <> "" And oblog.cacheConfig(5) = 1 Then
        Set rsUser = oblog.execute("select userid from oblog_user where user_domain='" & oblog.filt_badstr(user_domain) & "' and user_domainroot='" & oblog.filt_badstr(user_domainroot) & "' and userid<>" & userid)
        If Not rsUser.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>此域名已经被其他人使用！</li>"
        End If
    End If
	'进行数据有效性校验
	If oblog.cacheConfig(51)="1"  Then
		sMail=Trim(Request("postmail"))
		If  sMail<>"" Then
			if not oblog.IsValidEmail(sMail) then ErrMsg=ErrMsg & "<br><li>发布邮箱地址格式错误！</li>"
		End If
		sMobile=Trim(Request("postmobile"))
		If  sMobile<>"" Then
			If Len(sMobile) = 11 And IsNumeric(sMobile) Then
	        	If CInt(Left(sMobile, 3)) >= 134 And CInt(Left(sMobile, 3)) <= 139 Or CInt(Left(sMobile, 3)) = 159  Then
	            	'bMobile = True
	            Else
	            	ErrMsg=ErrMsg & "<br><li>您输入的手机号码错误或者系统暂不支持！</li>"
	            End If
	        Else
	        	ErrMsg=ErrMsg & "<br><li>您输入的手机号码错误或者系统暂不支持！</li>"
	        End If
	    'Else
	    '为空则不处理
		End If

		Dim rstMailPost
		Set rstMailPost=Server.CreateObject("adodb.recordset")

		'判断Mail是否重复
		If  sMail<>"" Then
			rstMailPost.open "select userid from oblog_user where postmail='" & LCase(Trim(sMail)) & "' And Userid<>" & UserID,conn,1,1
			If Not rstMailPost.Eof Then
				ErrMsg=ErrMsg & "<br><li>" & sMail & " 已经被使用,请更换发布邮箱!</li>"
			End If
			rstMailPost.Close
		End If
		'判断手机号码是否重复
		If  sMobile<>"" Then
			rstMailPost.open "select userid from oblog_user where postMobile='" & sMobile & "' And Userid<>" & UserID,conn,1,1
			If Not rstMailPost.Eof Then
				ErrMsg=ErrMsg & "<br><li>" &  sMobile & " 已经被使用,请更换发布号码!</li>"
			End If
			rstMailPost.Close
		End If
		If ErrMsg<>"" Then
			'Response.Write ErrMsg
			FoundErr=true
		End If
	End If

    If FoundErr = True Then
        Set rsUser = Nothing
        Exit Sub
    End If

    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from oblog_user where userID=" & userid
    If Not IsObject(Conn) Then link_database
    rsUser.open sqlUser, Conn, 1, 3
    If rsUser.bof And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>找不到指定的用户！</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If

	If API_Enable Then
		Dim blogAPI
		Set blogAPI = New DPO_API_OBLOG
		blogAPI.LoadXmlFile True
		blogAPI.UserName=rsUser("UserName")
		blogAPI.PassWord=Password
		blogAPI.EMail=Email
		blogAPI.Question=Question
		blogAPI.Answer=Answer
		blogAPI.Sex=Sex
		blogAPI.QQ=OICQ
		blogAPI.MSN=MSN
		blogAPI.userstatus=LockUser
		Call blogAPI.ProcessMultiPing("update")
		Set blogAPI=Nothing
	End If

    If Password <> "" Then
        rsUser("password") = md5(Password)
    End If
    rsUser("Question") = Question
    If Answer <> "" Then
        rsUser("Answer") = md5(Answer)
    End If
    rsUser("Sex") = Sex
    rsUser("userEmail") = Email
    If OICQ = "" Then
        OICQ = 0
    End If
    rsUser("qq") = OICQ
    rsUser("Msn") = MSN
    rsUser("User_Level") = User_Level
    rsUser("LockUser") = LockUser
    rsUser("user_isbest") = isbest
    rsUser("blogname") = blogname
    rsUser("user_classid") = usertype
    'rsUser("user_upfiles_max") = user_upfiles_max
    rsUser("user_upfiles_size") = upfiles_size
    rsUser("user_dir") = Trim(Request("user_dir"))
    If CheckAccess("r_user_group") Then
	    rsUser("user_group") = user_group
	    rsUser("scores") = scores
	End If
    If true_domain = 1 Then
        rsUser("custom_domain") = Trim(Request("custom_domain"))
    End If
    If Trim(Request("user_domain")) <> "" Then
        rsUser("user_domain") = Trim(Request("user_domain"))
    Else
        rsUser("user_domain") = " "
    End If
    rsUser("user_domainroot") = Trim(Request("user_domainroot"))
    If oblog.cacheConfig(51)="1"  Then
		rsuser("postmail")=Trim(Request("postmail"))
		rsuser("postmobile")=Trim(Request("postmobile"))
	End If
    rsUser.Update
    rsUser.Close
    Set rsUser = Nothing
	If User_Level = 6 Or LockUser = 1 Then
		oblog.Execute ("UPDATE oblog_log SET is_log_default_hidden = 1 WHERE userid in ("&userid&") or authorid in ("&userid&")")
	Else
		oblog.Execute ("UPDATE oblog_log SET is_log_default_hidden = 0 WHERE userid in ("&userid&") or authorid in ("&userid&")")
	End if
	Session ("CheckUserLogined_"&userName)=""
	WriteSysLog "进行了修改用户资料操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&Request.Form
    oblog.ShowMsg "修改成功!", ""
End Sub


Sub DelUser()
    Dim rs, i
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定要删除的用户</li>"
        Exit Sub
    End If
    If InStr(userid, ",") > 0 Then
        userid = Split(userid, ",")
        For i = 0 To UBound(userid)

			If API_Enable Then
				Dim struser,arruser,rstu
				Set rstu=oblog.execute("select username from oblog_user where userid="&userid(i))
				struser=rstu(0)
				arruser=struser&","&struser
			End If

            deloneuser (userid(i))
        Next

		If API_Enable Then
			Dim blogAPI
			Set blogAPI = New DPO_API_OBLOG
			blogAPI.LoadXmlFile True
			blogAPI.UserName=arruser
			Call blogAPI.ProcessMultiPing("delete")
			rstu.close
			Set rstu=Nothing
		End If

    Else

		If API_Enable Then
			Dim rst
			Set rst=oblog.execute("select username from oblog_user where userid="&userid)
			Set blogAPI = New DPO_API_OBLOG
			blogAPI.LoadXmlFile True
			blogAPI.UserName=rst(0)
			Call blogAPI.ProcessMultiPing("delete")
			rst.close
			Set rst=Nothing
			Set blogAPI=Nothing
		End If

		deloneuser (userid)
	End If
	If IsArray(userid) Then userid = Join(userid,",")
	WriteSysLog "进行了删除用户操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&OB_IIF(Request.QueryString,Request.Form)
'	Response.Redirect "m_user.asp"
    oblog.ShowMsg "删除成功!",""
End Sub

Sub deloneuser(userid)
    userid = CLng(userid)
    Dim rs, fso, f, uname, udir,rsfile,filename
    Set rs = oblog.execute("select user_dir,username,user_folder from oblog_user where userid=" & userid)
    If Not rs.EOF Then
        udir = rs(0)
        uname = rs(1)
        Set fso = Server.CreateObject(oblog.CacheCompont(1))
        if fso.FolderExists(Server.MapPath(blogdir & udir&"/"&rs("user_folder"))) then
            Set f = fso.GetFolder(Server.MapPath(blogdir & udir&"/"&rs("user_folder")))
            If Not IsNull(rs("user_folder")) Then f.Delete True
        End If
        Set f = Nothing
		'删除与日志相关联的一切数据库记录
		Dim blog
		Set blog=New Class_blog
		Call blog.DeleteFiles("",userid)
		Set blog=Nothing
        oblog.execute ("delete from oblog_log where userid=" & userid)
        oblog.execute ("delete from oblog_comment where userid=" & userid)
        oblog.execute ("delete from oblog_message where userid=" & userid)
        oblog.execute ("delete from oblog_subject where userid=" & userid)
        oblog.execute ("delete from oblog_user where userid=" & userid)
		'ATFLAG:连带删除用户已上传的所有文件
		Set rsfile=oblog.execute("select file_path from oblog_upfile where userid=" & userid)
		Do While Not rsfile.Eof
			filename=Server.Mappath(blogdir & rsfile(0))
'			Response.Write filename & "<BR>"
			If fso.FileExists(filename) Then fso.DeleteFile  filename,true
			rsfile.Movenext
		Loop
        oblog.execute ("delete from oblog_upfile where userid=" & userid)
        oblog.execute ("delete from oblog_friend where userid=" & userid)
        oblog.execute("update oblog_pm set dels=1 where sender='" &uname&"'")
		oblog.execute ("DELETE FROM oblog_album WHERE userid=" & userid)
    End If
	Set fso = Nothing
    Set rs = Nothing
End Sub

Sub LockUser()
    Dim rs
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请选择要锁定的用户</li>"
        Exit Sub
    End If
    userid = CLng(userid)
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select lockuser,is_log_default_hidden from oblog_user where userid=" & userid, Conn, 1, 3
    If Not rs.EOF Then
        rs(0) = 1
		rs(1) = 1
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
	oblog.Execute ("UPDATE oblog_log SET is_log_default_hidden = 1 WHERE userid = "&userid&" or authorid = "&userid)
	WriteSysLog "进行了锁定用户操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&Request.QueryString
    oblog.ShowMsg "锁定用户成功", ""
End Sub

Sub UnLockUser()
    Dim rs, udir
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请选择要锁定的用户</li>"
        Exit Sub
    End If
    userid = CLng(userid)
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select lockuser,istrouble,is_log_default_hidden from oblog_user where userid in (" & userid & ")", Conn, 1, 3
    If Not rs.EOF Then
        rs(0) = 0
		rs(1) = 0
		rs(2) = 0
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
	oblog.Execute ("UPDATE oblog_log SET is_log_default_hidden = 0 WHERE userid in ("&userid&") or authorid = "&userid)
	WriteSysLog "进行了解锁用户操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&Request.QueryString
    oblog.ShowMsg "解锁用户成功", ""
End Sub
Sub BlockUser()
    Dim rs
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请选择要前台屏蔽的用户</li>"
        Exit Sub
    End If
    userid = CLng(userid)
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select is_log_default_hidden from oblog_user where userid in (" & userid & ")", Conn, 1, 3
    If Not rs.EOF Then
        rs(0) = 1
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
	oblog.execute("update oblog_log set is_log_default_hidden = 1 where userid in (" & userid & ")")
	WriteSysLog "进行了屏蔽用户系统首页显示操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&Request.QueryString
    oblog.ShowMsg "屏蔽用户成功", ""
End Sub

Sub UnBlockUser()
    Dim rs, udir
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请选择要前台屏蔽的用户</li>"
        Exit Sub
    End If
    userid = CLng(userid)
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select is_log_default_hidden from oblog_user where userid in (" & userid & ")", Conn, 1, 3
    If Not rs.EOF Then
        rs(0) = 0
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
	oblog.execute("update oblog_log set is_log_default_hidden = 0 where userid in (" & userid & ")")
	WriteSysLog "进行了解锁用户系统首页显示操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&Request.QueryString
    oblog.ShowMsg "解锁用户成功", ""
End Sub
'审核
Sub PassIt(t)
    Dim rs, udir
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请选择要审核或取消审核的用户</li>"
        Exit Sub
    End If
	If Instr(userid,",") Then
		userid=FilterIds(userid)
	Else
		userid=Int(userid)
	End If
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select user_level from oblog_user where userid IN (" & userid &")", Conn, 1, 3
    If Not rs.EOF Then
		While Not rs.EOF
			rs(0) = t
			rs.Update
			rs.MoveNext
		Wend
    End If
    rs.Close
    Set rs = Nothing
    If t=6 Then
		oblog.Execute ("UPDATE oblog_log SET is_log_default_hidden = 1 WHERE userid in ("&userid&") or authorid in ("&userid&")")
		WriteSysLog "进行了取消用户审核操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&Request.QueryString
    	oblog.ShowMsg "取消用户审核成功", ""
    Else
		oblog.Execute ("UPDATE oblog_log SET is_log_default_hidden = 0 WHERE userid in ("&userid&") or authorid in ("&userid&")")
		WriteSysLog "进行了用户审核操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&Request.QueryString
    	oblog.ShowMsg "用户审核成功", ""
	End If
End Sub

Sub MoveUser()
	If Request.QueryString <>"" Then Exit Sub
    Dim msg,rst
    If userid = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定要移动的用户</li>"
        Exit Sub
    End If
    Dim User_Group
    User_Group = Trim(Request("User_Group"))
    If User_Group = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定目标用户组</li>"
        Exit Sub
    Else
        User_Group = CLng(User_Group)
    End If
    userid = FilterIds(userid)
    sql = "Update oblog_user set User_Group=" &User_Group &" where userID in (" & userid & ")"
    Conn.execute sql
    '重新计算组会员数目
    '重新进行计数
 	Set rst=oblog.Execute("select Count(UserId),user_group From oblog_user Where user_group>0 Group By user_group")
 	Do While Not rst.Eof
 		oblog.Execute("Update oblog_groups Set g_members=" & rst(0) & " Where groupid=" & rst(1))
 		rst.MoveNext
	Loop
	Set rst=Nothing
'	Response.Redirect "m_user.asp"
	WriteSysLog "进行了用户组转移操作，目标用户组ID："&User_Group&"",oblog.NowUrl&"?"&Request.Form
    oblog.ShowMsg "移动成功!",""
    'call WriteSuccessMsg(msg)
End Sub

Sub DoUpdate()
    Server.ScriptTimeOut = 999999999
    Dim BeginID, EndID, p1, rsUser, blog, i
	Dim Logintimes,B_Logs,E_Logs
	Dim sql
    BeginID = Trim(Request("BeginID"))
    EndID = Trim(Request("EndID"))
    Logintimes = Trim(Request("Logintimes"))
    B_Logs = Trim(Request("B_Logs"))
    E_Logs = Trim(Request("E_Logs"))
    If BeginID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定开始ID</li>"
    Else
        BeginID = CLng(BeginID)
    End If
    If EndID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定结束ID</li>"
    Else
        EndID = CLng(EndID)
    End If
    If Logintimes <> "" Then
		Logintimes = CLng(Logintimes)
    End If
    If B_Logs <> "" Then
		B_Logs = CLng(B_Logs)
	Else
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定开始ID</li>"
    End If
    If E_Logs <> "" Then
		E_Logs = CLng(E_Logs)
	Else
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定结束ID</li>"
    End If
	If B_Logs > E_Logs Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>结束日志ID需大于起始日志ID</li>"
	End IF
    If FoundErr = True Then Exit Sub
	If Logintimes > 0 Then
		sql = " AND Logintimes >"&Logintimes
	Else
		sql = ""
	End If
	If B_Logs> 0 And E_Logs > 0 Then
		sql = sql & " AND log_count >="&B_Logs&" AND log_count <="&E_Logs
	Else
		If B_Logs > 0 Then
			sql = sql & " AND log_count >="&B_Logs
		End If
		If E_Logs > 0 Then
			sql = sql & " AND log_count <="&E_Logs
		End if
	End If
	If Request("t0") <> 1 Then sql = ""
    Set rsUser = oblog.execute("select count(userid) from oblog_user where userID>=" & CLng(BeginID) & " "&sql&" and userID<=" & CLng(EndID))
    p1 = rsUser(0)
    Set rsUser = oblog.execute("select userid from oblog_user where userID>=" & CLng(BeginID) & " "&sql&" and userID<=" & CLng(EndID) & " order by userid")
    Set blog = New class_blog
    Response.Write ("<div style=""text-align: center;"">")
    Response.Write ("<div class=""progress1""><div class=""progress2"" id=""progress1""></div></div><span id=""pstr1""></span><br><br>")
    i = 1
    blog.progress_init
    Do While Not rsUser.EOF
		If Not IsObject(Conn) Then link_database
        Response.Write "<script>progress1.style.width =""" & Int(i / p1 * 100) & "%"";progress1.innerHTML=""" & Int(i / p1 * 100) & "%"";pstr1.innerHTML=""全部进度：当前用户ID:" & rsUser(0) & """;</script>" & vbCrLf
        Response.Flush
        blog.update_alllog_admin rsUser(0)
        rsUser.movenext
        i = i + 1
    Loop
    Response.Write ("</div>")
	'清除用户默认模板文件缓存
	blog.remove_user_skin_cache
    Set rsUser = Nothing
    Set blog = Nothing
	WriteSysLog "进行了更新用户静态页面操作，开始用户ID："&BeginID&"，结束用户ID："&EndID&"",oblog.NowUrl&"?"&Request.QueryString
End Sub

Sub DoUpdatelog()
    Server.ScriptTimeOut = 999999999
    Dim BeginID, EndID, p1, rs, blog, i
    BeginID = Trim(Request("BeginID"))
    EndID = Trim(Request("EndID"))
    If BeginID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定开始ID</li>"
    Else
        BeginID = CLng(BeginID)
    End If
    If EndID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定结束ID</li>"
    Else
        EndID = CLng(EndID)
    End If
    If FoundErr = True Then Exit Sub
    Set rs = oblog.execute("select count(logid) from oblog_log where logid>=" & CLng(BeginID) & " and logid<=" & CLng(EndID))
    p1 = rs(0)
    Set rs = oblog.execute("select logid,userid from oblog_log where logid>=" & CLng(BeginID) & " and logid<=" & CLng(EndID) & " order by logid")
    Set blog = New class_blog
    Response.Write ("<div style=""text-align: center;"">")
    Response.Write ("<div class=""progress1""><div class=""progress2"" id=""progress1""></div></div><span id=""pstr1""></span><br><br>")
    i = 1
    'blog.progress_init
    Do While Not rs.EOF
		If Not IsObject(Conn) Then link_database
        Response.Write "<script>progress1.style.width =""" & Int(i / p1 * 100) & "%"";progress1.innerHTML=""" & Int(i / p1 * 100) & "%"";pstr1.innerHTML=""进度：当前日志ID:" & rs(0) & """;</script>" & vbCrLf
        Response.Flush
        blog.userid = rs(1)
        blog.update_log rs(0), 0
        rs.movenext
        i = i + 1
    Loop
    Response.Write ("</div>")
    Set rs = Nothing
    Set blog = Nothing
	WriteSysLog "进行了更新日志静态页面操作，开始日志ID："&BeginID&"，结束日志ID："&EndID&"",oblog.NowUrl&"?"&Request.QueryString
End Sub
'重新计算用户日志、留言、评论及积分
Function ReScore()
	Dim userid
	userid=clng(Request("userid"))
	Dim rs,cmts,msgs,logs,bests,teams1,teams2,posts,scores,diggs
	Dim upfiles
	Set rs=oblog.Execute("select Count(*) From oblog_comment Where isdel=0 AND istate=1 and userid=" & userid)
	cmts=rs(0)
	Set rs=oblog.Execute("select Count(*) From oblog_Albumcomment Where isdel=0 AND istate=1 and userid=" & userid)
	cmts = cmts + rs(0)
	Set rs=oblog.Execute("select Count(*) From oblog_message Where isdel=0 AND istate=1 and  userid=" & userid)
	msgs=rs(0)
	Set rs=oblog.Execute("select Count(*) From oblog_log Where isdraft=0 and isdel=0 and  userid=" & userid)
	logs=rs(0)
	Set rs=oblog.Execute("select Count(*) From oblog_log Where isdraft=0 and passcheck=1 and isdel=0 and isbest=1 and  userid=" & userid)
	bests=rs(0)
	'申请数-
	Set rs=oblog.execute("select Count(*) From oblog_team Where createrid=" & userid)
	teams1=rs(0)
	'通过数目+
	Set rs=oblog.execute("select Count(*) From oblog_team Where createrid=" & userid)
	teams2=rs(0)
	'帖子积分
	Set rs=oblog.execute("select Count(*) From oblog_teampost Where logid>0 and  userid=" & userid)
	posts=rs(0)
	Set rs=oblog.execute("select Count(did) From oblog_digg Where authorid =" & userid&" AND diggtype=-1")
	diggs = rs(0)
	Set rs = oblog.Execute ("SELECT SUM(file_size) FROM oblog_upfile WHERE userid = "&userid)
	If Not rs.Eof Then upfiles = OB_IIF(rs(0),0) Else upfiles = 0
	If upfiles < 0 Then upfiles = 0
	Set rs=Nothing
	'注册初始分+日志分+精化分+回复分+留言分+群组帖子分+创建圈子奖励分-创建圈子消耗分
	Scores=Oblog.CacheScores(1)+Oblog.CacheScores(3)*logs+Oblog.CacheScores(5)*msgs+Oblog.CacheScores(6)*cmts+Oblog.CacheScores(7)*bests+Oblog.CacheScores(9)*teams2+Oblog.CacheScores(10)*posts-Oblog.CacheScores(8)*teams1+Oblog.CacheScores(22)*diggs
	oblog.Execute("Update Oblog_User Set log_count=" & logs & ",comment_count=" & cmts & ",message_count=" & msgs & ",scores=" & scores & ",user_upfiles_size = "&upfiles&",diggs="&diggs&" Where userid=" & userid)
	WriteSysLog "进行了用户资料修复操作，目标用户ID："&userid&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "用户资料（积分、日志数目等）修复完成！", ""
End Function
Set oblog = Nothing
%>
