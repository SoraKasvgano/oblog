<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_list_upfiles")=False Then Response.Write "无权操作":Response.End
dim rs, sql
dim userid,UserSearch,Keyword,strField
dim usermore,del
del=Trim(Request("del"))
userid=Trim(Request.QueryString("userid"))
'usermore=Trim(Request.QueryString("usermore"))
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
if usermore<>"" then
	G_P_FileName="m_uploadfile.asp?usermore=" & Usermore
else
	G_P_FileName="m_uploadfile.asp?UserSearch=" & UserSearch
end if
if Request("page")<>"" then
    G_P_This=cint(Request("page"))
else
	G_P_This=1
end if

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
		<li class="main_top_left left">上传文件管理(文件列表)</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">

<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="tdbg">
      <td width="100" height="30"><strong>管理导航：</strong></td>
      <td width="687" height="30"> &nbsp;&nbsp;&nbsp;&nbsp;<a href="m_uploadfile_user.asp">上传文件管理用户列表</a> | <a href="m_uploadfile.asp">上传文件管理文件列表</a></td>
    </tr>
    <form name="form2" method="post" action="m_uploadfile.asp">
  <tr class="tdbg">
      <td width="184">按文件名查询上传文件<strong>：</strong></td>
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
	'sGuide="<table width='98%' align='center'><tr><td align='left'>您现在的位置：<a href='m_uploadfile.asp'>上传文件管理-用户列表</a>&nbsp;&gt;&gt;&nbsp;"

	if Keyword="" then
		if usermore<>"" then
			sql="select top 500 file_name,file_path,file_size,isphoto,username,fileid from [oblog_upfile],oblog_user where oblog_upfile.userid=oblog_user.userid and oblog_upfile.userid="&usermore&" order by fileid desc"
			sGuide=sGuide & "用户id为"&usermore&"的用户上传文件"
		else
			sql="select top 500 file_name,file_path,file_size,isphoto,username,fileid from [oblog_upfile],oblog_user where oblog_upfile.userid=oblog_user.userid order by fileid desc"
			sGuide=sGuide & "前500个文件"
		end if
	else
		sql="select top 500 file_name,file_path,file_size,isphoto,username,fileid from [oblog_upfile],oblog_user where oblog_upfile.userid=oblog_user.userid and file_name like '%" & Keyword & "%' order by fileid  desc"
		sGuide=sGuide & "文件名中含有“ <font color=red>" & Keyword & "</font> ”的文件"
	end if

	sGuide=sGuide & "</td><td align='right'>"
	'Response.Write(sql)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		sGuide=sGuide & "共找到 <font color=red>0</font> 个上传文件</td></tr></table>"
	else
    	G_P_AllRecords=rs.recordcount
		sGuide=sGuide & "共找到 <font color=red>" & G_P_AllRecords & "</font> 个上传文件</td></tr></table>"
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
			if usermore<>"" then
			Response.Write oblog.showpage(true,true,"个上传文件")
			else
        	Response.Write oblog.showpage(true,true,"个用户")
			end if
   	 	else
   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
         	   	rs.move  (G_P_This-1)*G_P_PerMax
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
			    if usermore<>"" then
			    Response.Write oblog.showpage(true,true,"个上传文件")
			    else
            	Response.Write oblog.showpage(true,true,"个用户")
				end if
        	else
	        	G_P_This=1
           		showContent
				if usermore<>"" then
			    Response.Write oblog.showpage(true,true,"个上传文件")
			    else
           		Response.Write oblog.showpage(true,true,"个用户")
				end if
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
		<li class="main_top_left left"><%=sGuide%></li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <form name="myform" method="Post" action="m_uploadfile.asp?delmore=true" onsubmit="return confirm('确定要执行选定的操作吗？');">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
   <tr>
      <td style="border:0;margin:0;padding:0;">
<%do while not rs.EOF%>
<ul style="float:left;clear:none;width:140px;height: 220px; border:1px #efefef solid;margin:3px;padding:4px 0;">
<li style="text-align:center;">
    <a href="<%=blogdir & rs("file_path")%>" target="_blank" style="font-family:Arial,Helvetica,sans-serif;font-size:11px;">
	<%Dim sFileExt
    sFileExt=Right(Lcase(rs("file_path")),3)
    If sFileExt ="bmp" Or sFileExt="jpg" Or sFileExt="png" Or sFileExt="gif" Then
    Response.Write "<img src=""" & blogdir & rs("file_path") & """ width=""120"" height=""90"" border=""0"" />"
    End If
    %>
	<br />
    <%=rs("file_name")%></a>
</li>
<li style="padding:0 0 0 6px;">文件大小：<span style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#888;"><%=oblog.showsize(rs("file_size"))%></span></li>
<li style="padding:0 0 0 6px;">加入相册：<%if rs("isphoto")=1 then Response.Write("<span style=""color:#090;"">是</span>") else Response.Write("<span style=""color:#f00;"">否</span>")%></li>
<li style="padding:0 0 0 6px;">上传用户：<a href="../blog.asp?name=<%=rs("username")%>" target="_blank"><%=rs("username")%></a></li>
<li style="text-align:center;">
		<%
        Response.write "<input type=""checkbox"" id=""fileid"" name = ""fileid"" value="&rs("fileid")&" /><a href='m_uploadfile.asp?del="&rs("fileid")&"' onclick=""return confirm ('确认删除所选的文件？');""  style=""color:#f00;font-weight:600;"">删除文件</a>"
		%>
</li>
</ul>
          <%
		  	i=i+1
			if i>=G_P_PerMax then exit do

	rs.movenext
loop
%>

</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="140" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页所有内容</td>
            <td> <strong>操作：</strong>
              <input name="del" type="radio" value="del">
              删除&nbsp;&nbsp;
              &nbsp;&nbsp;
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

sub delfile(fid)
	Dim userid, filesize, filepath, fso, isphoto, imgsrc
	Dim fileid,fids,i
	fileid = FilterIDs (request("fileid"))
	If fid = "del" Then
		fid = fileid
	Else
		fid = CLng(fid)
	End If
	fids = Split (fid,",")
	For i = 0 To UBound(fids)
		sql="select userid ,file_size,file_path,isphoto from [oblog_upfile] where fileid=" & fids(i)
		Set rs = Server.CreateObject("adodb.recordset")
		rs.open sql, conn, 1, 3
		If Not rs.EOF Then
			userid = rs("userid")
			filesize = Int(rs("file_size"))
			filepath = blogdir & rs("file_path")
			isphoto = rs("isphoto")
			rs.Delete
			rs.Update
			rs.Close
			oblog.execute("update [oblog_user] set user_upfiles_num=user_upfiles_num-1,user_upfiles_size=user_upfiles_size-"&filesize&" where userid="&userid)
			If filepath <> "" Then
				imgsrc = filepath
				Set fso = Server.CreateObject(oblog.CacheCompont(1))
				If InStr("jpg,bmp,gif,png,pcx", Right(imgsrc, 3)) > 0 Then
					imgsrc = Replace(imgsrc, Right(imgsrc, 3), "jpg")
					imgsrc = Replace(imgsrc, Right(imgsrc, Len(imgsrc) - InStrRev(imgsrc, "/")), "pre" & Right(imgsrc, Len(imgsrc) - InStrRev(imgsrc, "/")))
					If fso.FileExists(Server.MapPath(imgsrc)) Then
						fso.DeleteFile Server.MapPath(imgsrc)
					End If
				End If
				If fso.FileExists(Server.MapPath(filepath)) Then
					fso.DeleteFile Server.MapPath(filepath)
				End If
				Set fso = Nothing
			End If
			If isphoto = 1 Then
				Set rs = oblog.Execute ("SELECT COUNT(commentid) FROM oblog_albumcomment WHERE mainid="&fids(i))
				oblog.execute ("update [oblog_user] set comment_count = comment_count -"&OB_IIF(rs(0),0)&" where userid="&userid)
				rs.close
				oblog.Execute ("DELETE FROM oblog_album WHERE fileid = "&fids(i))
				oblog.execute ("DELETE FROM [oblog_albumcomment] WHERE mainid = "&fids(i))
			End if
		Else
			rs.Close
		End If
	Next
	Set rs = Nothing
	WriteSysLog "进行了删除上传文件操作，目标文件ID："&fid&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "删除成功！",""
end sub
Set oblog = Nothing
%>
</body>
</html>