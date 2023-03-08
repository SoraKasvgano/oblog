<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_user_upfiles")=False Then Response.Write "无权操作":Response.End
dim rs, sql
dim userid,UserSearch,Keyword,strField
dim uppath,fso,thefile
dim del,moreid,delmore,rstGroup
moreid=Trim(Request("moreid"))
'Response.Write moreid
del=Trim(Request.QueryString("del"))
userid=Trim(Request.QueryString("userid"))
delmore=Trim(Request.QueryString("delmore"))
keyword=Trim(Request("keyword"))
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
UserSearch=Trim(Request("UserSearch"))
Action=Trim(Request("Action"))
if UserSearch="" then
	UserSearch=0
else
	UserSearch=CLng(UserSearch)
end if
G_P_FileName="m_uploadfile_user.asp?UserSearch=" & UserSearch
if Request("page")<>"" then
    G_P_This=cint(Request("page"))
else
	G_P_This=1
end if
Set rstGroup=Server.CreateObject("Adodb.Recordset")
rstGroup.Open "select groupid,g_up_space From Oblog_groups",conn,1,3
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
		<li class="main_top_left left">上传文件管理(用户列表)</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="tdbg">
      <td width="100" height="30"><strong>管理导航：</strong></td>
      <td width="687" height="30"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="m_uploadfile_user.asp">上传文件管理用户列表</a> <a href="m_uploadfile.asp">上传文件管理文件列表</a></td>
    </tr>
	<form name="form2" method="post" action="m_uploadfile_user.asp">
  <tr class="tdbg">
    <td width="184">按用户查询上传文件：</td>
    <td width="236">
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10">
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
if del<>"" then
	Delfile del
else
	call main()
end if

sub main()
	dim theFolder,filecount,totalsize,upstr
	upstr=" where user_upfiles_size>0"
	sGuide="<table width='98%' align='center'><tr><td align='left'>您现在的位置：<a href='m_Uploadfile_user.asp'>上传文件管理-用户列表</a>&nbsp;&gt;&gt;&nbsp;"

	if Keyword="" then
		sql="select top 500 user_upfiles_size,username,userid,user_group,user_upfiles_num from [oblog_user] "&upstr&" order by user_upfiles_size desc"
		sGuide=sGuide & "前500个用户"
	else
		sql="select top 500 user_upfiles_size,username,userid,user_group,user_upfiles_num from [oblog_user] "&upstr&" and userName like '%" & Keyword & "%' order by user_upfiles_size  desc"
		sGuide=sGuide & "用户名中含有“ <font color=red>" & Keyword & "</font> ”的用户"
	end if

	sGuide=sGuide & "</td><td align='right'>"
	'Response.Write(sql)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		sGuide=sGuide & "共找到 <font color=red>0</font> 个有上传文件的用户</td></tr></table>"
		Response.write sGuide
	else
    	G_P_AllRecords=rs.recordcount
		sGuide=sGuide & "共找到 <font color=red>" & G_P_AllRecords & "</font> 个有上传文件的用户</td></tr></table>"
		Response.write sGuide
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
        	Response.Write oblog.showpage(true,true,"个用户")
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rs.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	Response.Write oblog.showpage(true,true,"个用户")
        	else
	        	G_P_This=1
           		showContent
           		Response.Write oblog.showpage(true,true,"个用户")
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
end sub

sub showContent()
	dim i
	dim user_maxsize,vip_maxsize,m_maxsize,umix,uleft,uimg
    i=0
	'user_maxsize=oblog.setup(36,0)
	'vip_maxsize=oblog.setup(40,0)
	'm_maxsize=oblog.setup(44,0)
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">上传文件管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="myform" method="Post" action="m_uploadfile_user.asp?delmore=true" onsubmit="return confirm('确定要执行选定的操作吗？');">
<style type="text/css">
<!--
.border tr td {padding:3px 5px!important;}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title">
            <td width="110" align="center"><strong>用户</strong></td>
            <td width="80" align="center"><strong>上传文件数</strong></td>
            <td width="90" align="center"><strong>总计大小</strong></td>
            <td width="90" align="center"><strong>剩余空间</strong></td>
            <td width="90" align="center"><strong>分配空间</strong></td>
            <td align="center"><strong>百分比</strong></td>
            <td width="70" align="center"><strong>操作</strong></td>
          </tr>
<%do while not rs.EOF
	rstGroup.Filter="groupid=" & rs("user_group")
	If Not rstGroup.Eof Then
		umix=rstGroup(1)
		If umix="" Or umix=0 Then
			umix="不限制"
			uleft="不限"
			uimg=0
		Else
			uimg=((rs("user_upfiles_size")/1024)/umix)*100
			uleft=oblog.showSize((umix*1024-rs("user_upfiles_size")))
			umix=oblog.showSize(umix*1024)
		End If
	End If

%>
          <tr class="tdbg">
            <td><a href="../blog.asp?name=<%=rs("username")%>" target="_blank"><%=rs("username")%></a></td>
            <td align="center" style="font-family:Arial,Helvetica,sans-serif;font-size:10px;font-weight:600;"><%=rs("user_upfiles_num")%></td>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#f00;"><%=oblog.ShowSize(rs("user_upfiles_size"))%></td>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#090;"><%=uleft%></td>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#217dbd;"><%=umix%> </td>
            <td align="center"><div align="left"><img src="images/bar.gif" width=<%=uimg&"%"%> height=10></div></td>
            <td align="center"><%
        Response.write "<a href='m_uploadfile.asp?usermore="&rs("userid")&"'>详细</a>&nbsp;"
        Response.write "<a href='m_uploadfile_user.asp?del="&rs("userid")&"'>清空</a>"

		%> </td>
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

sub delfile(uid)
	Server.ScriptTimeOut=999999999
	uid=CLng(uid)
	dim rs,fs,oFolder,file,album_commentnum
	set rs=oblog.execute("select file_path from oblog_upfile where userid="&Int(uid))
	if not rs.eof then
		set fs=CreateObject(oblog.CacheCompont(1))
		Do While Not rs.Eof
			On Error Resume Next
			fs.DeleteFile(Server.mappath(blogdir& rs(0)))
			rs.Movenext
		Loop
		Set rs =oblog.Execute ("SELECT COUNT(commentid) FROM oblog_albumcomment WHERE userid="&Int(uid))
		album_commentnum = RS(0)
		If IsNull(album_commentnum) Then album_commentnum = 0
		set rs=nothing
		set fs=nothing
		oblog.execute("delete from [oblog_upfile] where userid="&uid)
		oblog.execute("delete from [oblog_album] where userid="&uid)
		oblog.execute("delete from [oblog_albumcomment] where userid="&uid)
		oblog.execute("update [oblog_user] set user_upfiles_size=0,user_upfiles_num=0,comment_count = comment_count -"&album_commentnum&" where userid="&uid)
	end If
	WriteSysLog "进行了清空用户上传文件操作，目标用户ID："&uid&"",oblog.NowUrl&"?"&Request.QueryString
	Response.Redirect("m_uploadfile_user.asp")
end Sub
Set oblog = Nothing
%>
</body>
</html>