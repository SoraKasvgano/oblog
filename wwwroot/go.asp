<!-- #include file="inc/inc_syssite.asp" -->
<%
dim logid,commentid,userid,messageid,action,user,sThisMonth,sMonth,teamid,albumid,Fileid
dim rs,uf,sql
Dim goUrl
logid=Trim(Request("logid"))
commentid=Trim(Request("commentid"))
messageid=Trim(Request("messageid"))
userid=Trim(Request("userid"))
user=Trim(Request("user"))
action=Trim(Request("action"))
teamid=Trim(Request("teamid"))
albumid=Trim(Request("albumid"))
Fileid=Trim(Request("fileid"))
goUrl=Trim(request("url"))
If logid<>"" And IsNumeric(logid) Then
	logid=CLng(logid)
Else
	logid=""
End If
If commentid<>"" And IsNumeric(commentid) Then
	commentid=CLng(commentid)
Else
	commentid=""
End If
If messageid<>"" And IsNumeric(messageid) Then
	messageid=CLng (messageid)
Else
	messageid=""
End If
If userid<>"" And IsNumeric(userid) Then
	userid=CLng (userid)
Else
	userid=""
End If
If teamid<>"" And IsNumeric(teamid) Then
	teamid=CLng (teamid)
Else
	teamid=""
End If
If albumid<>"" And IsNumeric(albumid) Then
	albumid=CLng (albumid)
Else
	albumid=""
End If
If Fileid<>"" And IsNumeric(Fileid) Then
	Fileid=CLng (Fileid)
Else
	Fileid=""
End If
if not IsObject(conn) then link_database
if logid<>"" then
	if action="up" then
		sql="select top 1 logfile from oblog_log where logid<"&logid&" and userid="&userid&" order by logid desc"
	elseif action="down" then
		sql="select top 1 logfile from oblog_log where logid>"&logid&" and userid="&userid&" order by logid"
	else
		sql="select logfile from oblog_log where logid="&logid
	end if
	set rs=conn.execute(sql)
	if not rs.eof then
		dim logf
		logf=rs(0)
		set rs=nothing
		if isobject(conn) then conn.close:set conn=Nothing
		If logf = "" Or IsNull(logf) Then Response.Write("此日志为草稿，请登录用户草稿箱查看内容"):Response.End
		Response.Redirect(logf)
	else
		set rs=nothing
		Response.Write("无此日志")
	end if
elseif userid<>"" then
	set rs=conn.execute("select user_dir,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_user where userid="&CLng(userid))
	if not rs.eof then
		if true_domain=1 then
			if rs("custom_domain")="" or isnull(rs("custom_domain")) then
				uf="http://"&rs("user_domain")&"."&rs("user_domainroot")&"/index."&f_ext
			else
				uf="http://"&rs("custom_domain")&"/index."&f_ext
			end if
		else
			uf=rs("user_dir")&"/"&rs("user_folder")&"/index."&f_ext
		end if
		set rs=nothing
		if isobject(conn) then conn.close:set conn=nothing
		Response.Redirect(uf)
	else
		set rs=nothing
		Response.Write("无此用户")
	end if
elseif messageid<>"" then
	set rs=conn.execute("select messagefile,user_dir,oblog_user.userid,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_message,oblog_user where oblog_message.userid=oblog_user.userid and messageid="&CLng (messageid))

	if not rs.eof then
		if true_domain=1 then
			if rs("custom_domain")="" or isnull(rs("custom_domain")) then
				uf="http://"&rs("user_domain")&"."&rs("user_domainroot")&"/message."&f_ext
			else
				uf="http://"&rs("custom_domain")&"/message."&f_ext
			end if
		else
			uf=rs("user_dir")&"/"&rs("user_folder")&"/message."&f_ext
		end if
		set rs=nothing
		if isobject(conn) then conn.close:set conn=nothing
		Response.Redirect(uf)
	else
		set rs=nothing
		Response.Write("无此留言")
	end if
Elseif albumid<>"" Then
	set rs=conn.execute("select user_dir,user_folder,user_domain,user_domainroot"&str_domain&" ,userid from oblog_user where userid="&CLng (albumid))
	if not rs.eof then
		if true_domain=1 then
			if rs("custom_domain")="" or isnull(rs("custom_domain")) then
				uf="http://"&rs("user_domain")&"."&rs("user_domainroot")&"/cmd." & f_ext & "?uid=" & rs("userid") & "&do=album"
			else
				uf="http://"&rs("custom_domain")&"/cmd." & f_ext & "?uid=" & rs("userid") & "&do=album"
			end if
		else
			uf=rs("user_dir")&"/"&rs("user_folder")&"/cmd." & f_ext & "?uid=" & rs("userid") & "&do=album"
		end if
		set rs=nothing
		if isobject(conn) then conn.close:set conn=nothing
		Response.Redirect(uf)
	else
		set rs=nothing
		Response.Write("无此用户")
	end if
'按用户名访问
elseif user<>"" Then
	user=Replace(user,"'","")
	user=Replace(user,"%","")
	user=Replace(user," ","")
	user=Replace(user,"--","")
	If user<>"" Then
		set rs=conn.execute("select user_dir,user_folder From oblog_user where username='" & user & "'")
		if not rs.eof then
			uf=rs(0)&"/"&rs(1)&"/index."&f_ext
			set rs=nothing
			if isobject(conn) then conn.close:set conn=nothing
			Response.Redirect(uf)
		else
			set rs=nothing
			Response.Write("无此用户")
		end if
	Else
		Response.Write("无此用户")
	End If
ElseIf Fileid<>"" Then
	Set rs = conn.Execute ("select a.userid,user_dir,user_folder,teamid,user_domain,user_domainroot"&str_domain&" FROM oblog_user a,oblog_album b WHERE b.fileid = "&Fileid&" AND a.userid =b.userid")
	If Not rs.EOF Then
		If rs("teamid") = 0 Then
			if true_domain=1 then
				if rs("custom_domain")="" or isnull(rs("custom_domain")) then
					uf="http://"&rs("user_domain")&"."&rs("user_domainroot")&"/cmd."&f_ext & "?do=photocomment&fileid="&fileid&"&uid="&rs(0)
				else
					uf="http://"&rs("custom_domain")&"/cmd."&f_ext & "?do=photocomment&fileid="&fileid&"&uid="&rs(0)
				end if
			else
				uf=rs("user_dir")&"/"&rs("user_folder")&"/cmd."&f_ext & "?do=photocomment&fileid="&fileid&"&uid="&rs(0)
			end if
			set rs=Nothing
		Else
			uf = "group.asp?cmd=photocomment&gid="&rs("teamid")&"&fileID="&Fileid
		End if
		if isobject(conn) then conn.close:set conn=nothing
		Response.Redirect(uf)
	Else
		Response.Write "无此记录"
	End If
ElseIf goUrl <>"" Then
	If oblog.chk_badword(goUrl) Then
		Response.Write "地址非法"
		Response.End
	Else
		goUrl = oblog.filt_badword(goUrl)
		Response.Redirect goUrl
	End if
Else
	Response.Write "参数错误"
end if
if IsObject(conn) then conn.close:set conn=nothing
%>