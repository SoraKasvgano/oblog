<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_blogstar")=False Then Response.Write "无权操作":Response.End
dim rs, sql
dim id,UserSearch,Keyword,strField,douname
douname=oblog.filt_badstr(request("douname"))

keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
strField=Trim(Request("Field"))
UserSearch=Trim(Request("UserSearch"))
Action=Trim(Request("Action"))
id=Trim(Request("id"))

if UserSearch="" then
	UserSearch=0
else
	UserSearch=CLng(UserSearch)
end if
G_P_FileName="m_blogstar.asp?UserSearch=" & UserSearch
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
		<li class="main_top_left left">博 客 之 星 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form name="form1" action="m_blogstar.asp" method="get">
    <tr class="tdbg">
      <td width="100" height="30"><strong>快速查找：</strong></td>
      <td width="687" height="30"><select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>请选择查询条件</option>
		  <option value="0">最后500个博客之星</option>
          <option value="1">通过审核的博客之星</option>
          <option value="2">未通过审核的博客之星</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="m_blogstar.asp">博客之星管理首页</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="m_blogstar.asp">
  <tr class="tdbg">
      <td width="120"><strong>高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">
	  <option value="blogname" selected>博客之星名</option>
	  <option value="username" selected>用户名</option>
	  <option value="nickname" selected>用户昵称</option>
      <option value="UserID" >博客之星ID</option>

      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10">
        若为空，则查询所有</td>
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
select Case LCase(Action)
	Case "modify"
		call Modify()
	Case "savemodify"
		call SaveModify()
	Case "del"
		call DelUser()
	Case "pass0"
		Call Pass(0)
	Case "pass1"
		Call Pass(1)
	Case else
		call main()
end select
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	dim G_P_Guide
	G_P_Guide="<table width='100%'><tr><td align='left'>您现在的位置：<a href='m_blogstar.asp'>博客之星管理</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from oblog_blogstar order by id desc"
			G_P_Guide=G_P_Guide & "最后500个博客之星"
		case 1
			sql="select * from oblog_blogstar where ispass=1 order by id desc"
			G_P_Guide=G_P_Guide & "通过审核的博客之星"
		case 2
			sql="select * from oblog_blogstar where ispass=0 order by id desc"
			G_P_Guide=G_P_Guide & "未通过审核的博客之星"
		case 10
			if Keyword="" then
				sql="select top 500 * from oblog_blogstar order by id desc"
				G_P_Guide=G_P_Guide & "所有博客之星"
			else
				select case strField
				case "UserID"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID必须是整数！</li>"
					else
						sql="select * from oblog_blogstar where id =" & CLng(Keyword)
						G_P_Guide=G_P_Guide & "博客之星ID等于<font color=red> " & CLng(Keyword) & " </font>的博客之星"
					end if
				case "blogname"
					sql="select * from oblog_blogstar where blogname like '%" & Keyword & "%' order by id  desc"
					G_P_Guide=G_P_Guide & "博客名中含有“ <font color=red>" & Keyword & "</font> ”的博客之星"
				case "username"
					sql="select * from oblog_blogstar where username like '%" & Keyword & "%' order by id  desc"
					G_P_Guide=G_P_Guide & "用户名中含有“ <font color=red>" & Keyword & "</font> ”的博客之星"
				case "nickname"
					sql="select * from oblog_blogstar where usernickname like '%" & Keyword & "%' order by id  desc"
					G_P_Guide=G_P_Guide & "博客名中含有“ <font color=red>" & Keyword & "</font> ”的博客之星"
				end select
			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>错误的参数！</li>"
	end select
	G_P_Guide=G_P_Guide & "</td><td align='right'>"
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		G_P_Guide=G_P_Guide & "共找到 <font color=red>0</font> 个博客之星</td></tr></table>"
		Response.write G_P_Guide
	else
    	G_P_AllRecords=rs.recordcount
		G_P_Guide=G_P_Guide & "共找到 <font color=red>" & G_P_AllRecords & "</font> 个博客之星</td></tr></table>"
		Response.write G_P_Guide
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
        	Response.write oblog.showpage(true,true,"个博客之星")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rs.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	Response.write oblog.showpage(true,true,"个博客之星")
        	else
	        	G_P_This=1
           		showContent
           		Response.write oblog.showpage(true,true,"个博客之星")
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i
    i=0
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">博 客 之 星 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="myform" method="Post" action="m_blogstar.asp" onsubmit="return confirm('确定要执行选定的操作吗？');">
<style type="text/css">
<!--
.border tr td {padding:3px 0!important;}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="title">
    <td align="center" width="44"><strong>ID</strong></td>
    <td align="center" width="100"><strong>博客之星图片</strong></td>
    <td align="center" width="120"><strong>申请博客 申请时间</strong></td>
    <td align="center"><strong>博客之星简介</strong></td>
	  <td align="center" width="90"><strong>用户名/昵称</strong></td>
    <td align="center" width="70"><strong>审核操作</strong></td>
    <td align="center" width="70"><strong>管理操作</strong></td>
  </tr>
          <%do while not rs.EOF %>
  <tr class="tdbg">
    <td align="center" style="font-family:Century Gothic,verdana,tahoma,Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("id")%></td>
    <td align="center">
	<a href="<%=rs("picurl")%>" target="_blank" title="点击查看该图"><img src="<%=ProIco(rs("picurl"),1)%>" align="absmiddle" style="width:80px;height:60px;border:0;"></a>
	</td>
    <td>
	<span style="display:block;color:#666;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:12px;padding:0 0 0 8px!important;"><a href="<%=rs("userurl")%>" target="_blank" title="点击访问该博客"><%=rs("blogname")%></a></span>
	<span style="display:block;color:#999;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:10px;padding:0 0 0 8px!important;">
	<%
	if rs("addtime")<>"" then
		Response.write rs("addtime")
	else
		Response.write "&nbsp;"
	end if
	%>
	</span>
	</td>
    <td valign="top"><span style="font-family:tahoma,Arial,Helvetica,sans-serif;padding:0 4px 0 4px!important;"><%=oblog.filt_html(rs("info"))%></span></td>
	<td align="center"><a href="<%=rs("userurl")%>" target="_blank" title="点击访问该博客"><%=rs("username")&"<br/>"&rs("usernickname")%></a></td>
    <td align="center">
	<%
	select case rs("ispass")
		case 0
			Response.write "<span style=""color:#f30;font-weight:600;"">待审</span>"
		case 1
			Response.write "<span style=""color:#090;font-weight:600;"">通过</span>"
	end select
	%>&nbsp;
	<%
	If  rs("ispass")=0 Then
		Response.write "<a href='m_blogstar.asp?Action=pass1&id=" & rs("id") & "&douname="&rs("username")&"'>通过</a>&nbsp;"
	Else
		Response.write "<a href='m_blogstar.asp?Action=pass0&id=" & rs("id") & "'>取消</a>&nbsp;"
	End If
	%>
	</td>
    <td align="center">
<%
Response.write "<a href='m_blogstar.asp?Action=Modify&id=" & rs("id") & "'>修改</a>&nbsp;"
Response.write "<a href='m_blogstar.asp?Action=Del&id=" & rs("id") & "' onClick='return confirm(""确定要删除此博客之星吗？"");'>删除</a>&nbsp;"
%>
	</td>
  </tr>
          <%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
%>
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


sub Modify()
	dim rsUser,sqlUser
	id=CLng(id)
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_blogstar where id=" & id
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的博客之星！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">修改博客之星信息</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<FORM name="Form1" action="m_blogstar.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td>blog名：</td>
      <td><input name=blogname   type=text id="blogname" value="<%=rsuser("blogname")%>" size=30 maxlength=20></td>
    </tr>
    <TR class="tdbg" >
      <TD width="40%">连接地址：</TD>
      <TD width="60%"> <INPUT name="userurl" value="<%=rsUser("userurl")%>" size=50   maxLength=250> <a href="<%=rsuser("userurl")%>" target="_blank">查看</a>
      </TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%"> 图片连接(<strong><font color="#FF0000">请将大图片手工改为合适的尺寸</font></strong>)：</TD>
      <TD width="60%"> <INPUT name=picurl value="<%=rsUser("picurl")%>" size=50 maxLength=250><a href="<%=rsuser("picurl")%>" target="_blank">查看</a></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">简介：</TD>
      <TD width="60%"><textarea name="bloginfo" cols="40" rows="5"><%=oblog.filt_html(rsuser("info"))%></textarea></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="40%">状态：</TD>
      <TD width="60%"><input type="radio" name="ispass" value=0 <%if rsUser("ispass")=0 then Response.write "checked"%>>
        未通过审核&nbsp;&nbsp; <input type="radio" name="ispass" value=1 <%if rsUser("ispass")=1 then Response.write "checked"%>>
        已通过审核</TD>
    </TR>
    <TR class="tdbg" >
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="保存修改结果"> <input name="id" type="hidden" id="id" value="<%=rsUser("id")%>"></TD>
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
	rsUser.close
	set rsUser=nothing
end sub


sub SaveModify()
	If Request.QueryString <>"" Then Exit Sub
	dim rsuser,sqlUser
	id=CLng(id)
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_blogstar where id=" & id
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	rsUser("blogname")=Trim(Request("blogname"))
	rsUser("userurl")=Trim(Request("userurl"))
	rsUser("picurl")=Trim(Request("picurl"))
	rsUser("ispass")=Trim(Request("ispass"))
	rsUser("info")=Trim(Request("bloginfo"))
	rsUser("addtime")=oblog.ServerDate(now())
	rsUser.update
	rsUser.Close
	set rsUser=Nothing
	WriteSysLog "进行了修改博客之星资料操作，目标用户ID："&id&"",""
	oblog.ShowMsg "修改成功!",""
end sub

sub DelUser()
	id=CLng(id)
	oblog.execute("delete from oblog_blogstar where id="&id)
	WriteSysLog "进行了删除博客之星操作，目标用户ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "删除成功！",""
end sub

sub Pass(iState)
	id=CLng(id)
	oblog.execute("Update  oblog_blogstar Set ispass="& Cint(iState) &" where id="&id)
	If iState=0 Then
		WriteSysLog "进行了取消博客之星操作，目标用户ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		oblog.ShowMsg "已取消该博客之星资格！",""
	Else
		WriteSysLog "进行了批准博客之星操作，目标用户ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
		If int(oblog.CacheConfig(86)) = 1 Then 
		oblog.execute("INSERT INTO oblog_pm(incept,sender,topic,content) VALUES('"&doUname&"','系统管理员','系统通知!您成为本站博客之星!','恭喜,您已经被批准成为本站光荣的博客之星!再接再励哦!(此信息系统自动发出,阅读后将被自动删除.您不必回复!)')")
		End If 
		oblog.ShowMsg "已批准该博客之星申请！",""
	End If
end Sub
Set oblog = Nothing
%>