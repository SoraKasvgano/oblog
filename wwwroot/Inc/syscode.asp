<%
Dim G_P_Show
G_P_Show=G_P_Show& "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"& vbcrlf
G_P_Show=G_P_Show& "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbcrlf
G_P_Show=G_P_Show& "<head>" & vbcrlf
G_P_Show=G_P_Show& "<meta http-equiv=""content-type"" content=""text/html; charset=gb2312""/>" & vbcrlf
G_P_Show=G_P_Show& "<meta name=""generator"" content=""oblog""/>" & vbcrlf
G_P_Show=G_P_Show& "<meta name=""keywords"" content="""&oblog.cacheConfig(9)&"""/>" & vbcrlf
G_P_Show=G_P_Show& "<link rel=""alternate"" href=""rssfeed.asp"" type=""application/rss+xml"" title=""$show_title$的频道文章列表"" />" & vbcrlf
G_P_Show=G_P_Show& "<title>$show_title$</title>" & vbcrlf
G_P_Show=G_P_Show& "<link href=""OblogStyle/OblogsysDefault4.css"" rel=""stylesheet"" type=""text/css"" /> " & vbcrlf
G_P_Show=G_P_Show& "<script src=""inc/main.js""></script>" & vbcrlf
G_P_Show=G_P_Show& "<script>function chkdiv(divid){var chkid=document.getElementById(divid);if(chkid != null){return true; }else {return false; }}</script>" & vbcrlf
G_P_Show=G_P_Show& "{OB_STYLE}" & vbcrlf
G_P_Show=G_P_Show& "</head>" & vbcrlf
G_P_Show=G_P_Show& "<body>" & vbcrlf

Function show_userlogin(n)
	if n=0 then
		show_userlogin = "<div id=""ob_login""></div>"&"<script src=""login.asp?action=showindexlogin""></script>"
	else
		show_userlogin = "<div id=""ob_login""></div>"&"<script src=""login.asp?action=showindexlogin&n=1""></script><script src=""inc/main.js""></script>"
	end if
End Function

Sub indexshow()
	G_P_Show = Replace (G_P_Show,"$show_title$",oblog.cacheConfig(2))

	If InStr(G_P_Show, "$show_sitename$") > 0 Then
		G_P_Show = Replace(G_P_Show, "$show_sitename$", oblog.cacheconfig(1))
	End If

	If InStr(G_P_Show, "$show_placard$") > 0 Then
		G_P_Show = Replace(G_P_Show, "$show_placard$", show_placard())
	End If

	If InStr(G_P_Show, "$show_friends$") > 0 Then
	G_P_Show = Replace(G_P_Show, "$show_friends$", show_friends())
	End If

	If InStr(G_P_Show, "$show_count$") > 0 Then
		G_P_Show = Replace(G_P_Show, "$show_count$", show_count())
	End If

	If InStr(G_P_Show, "$show_userlogin$") > 0 Then
		G_P_Show = Replace(G_P_Show, "$show_userlogin$", show_userlogin(0))
	End If

	If InStr(G_P_Show, "$show_userlogin_l$") > 0 Then
		G_P_Show = Replace(G_P_Show, "$show_userlogin_l$", show_userlogin(1))
	End If

	If InStr(G_P_Show, "$show_xml$") > 0 Then
		G_P_Show = Replace(G_P_Show, "$show_xml$", show_sysxml())
	End If

	If InStr(G_P_Show, "$show_blogstar$") > 0 Then
		G_P_Show = Replace(G_P_Show, "$show_blogstar$", show_blogstar())
	End If

	If InStr(G_P_Show, "$show_cloudtags$") > 0 Then
		G_P_Show = Replace(G_P_Show, "$show_cloudtags$", Tags_SystemTags("1"))
	End If

	Call runsub("$show_newblogger")
	Call runsub("$show_comment")
	Call runsub("$show_subject")
	Call runsub("$show_blogupdate")
	Call runsub("$show_bestblog")
	Call runsub("$show_bloger")
	Call runsub("$show_class")
	Call runsub("$show_log")
	'
	Call runsub("$show_zt")
	Call runsub("$show_m")
	'
	Call runsub("$show_userlog")
	Call runsub("$show_search")
	Call runsub("$show_cityblogger")
	Call runsub("$show_newphoto")
	Call runsub("$show_blogstar2")
	Call runsub("$show_teams")
	Call runsub("$show_posts")
	Call runsub("$show_hotblog")
	Call runsub("$show_hottag")
	Call runsub("$show_treeclass")
	Call runsub("$show_bl")
	Call runsub("$show_template")
	Call runsub("$show_album")
	Call runsub("$show_pic")
	Call runsub("$show_diggs")
	Call runsub("$show_userdiggs")
	Call runsub("$show_rnduser")
	Call runsub("$show_indexlog")
End Sub

Sub sysshow()
	if Application(oblog.cache_name&"_list_update")=False And application(oblog.cache_name&"list")<>"" Then
		G_P_Show=application(oblog.cache_name&"list")
	Else
		Dim rstmp,sContent,sStyle
		Set rstmp = oblog.execute("select skinshowlog from oblog_sysskin where isdefault=1")
		sContent=rstmp(0)
		sStyle=OB_PickUpCss(sContent)
		G_P_Show=Replace(G_P_Show,"{OB_STYLE}",sStyle)
		G_P_Show = Replace (G_P_Show,"$show_title$","$show_title_list$")
		'Response.Write sStyle
		'Response.Write sContent
		G_P_Show=G_P_Show&sContent
		Set rstmp = Nothing
		'副模板取消城市选项
		G_P_Show=Replace(G_P_Show,"show_cityblogger(0)$","")
		G_P_Show=Replace(G_P_Show,"show_cityblogger(1)$","")
		Call indexshow
		Application.Lock
		application(oblog.cache_name&"_list_update")=False
		application(oblog.cache_name&"list")=G_P_Show
		Application.unLock
	End If
End Sub

Sub runsub(label)
	On Error Resume Next
	Dim tmp1, tmp2, i
	Dim tmpstr, para
	tmp2 = 1
	While InStr(tmp2, G_P_Show, label) > 0
		tmp1 = InStr(tmp2, G_P_Show, label)
		tmp2 = InStr(tmp1 + 1, G_P_Show, "$")
		tmpstr = Mid(G_P_Show, tmp1, tmp2 - tmp1)
		tmpstr = Replace(tmpstr, "(", "")
		tmpstr = Replace(tmpstr, ")", "")
		tmpstr = Trim(Replace(tmpstr, label, ""))
		para = Split(tmpstr, ",")
		select Case label
		Case "$show_log"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_log(para(0),para(1),para(2),para(3),para(4),para(5),para(6),para(7),para(8)))
			If Err Then
				Response.Write "<br/>$show_log$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		'''''''''''''''''''''''''''''
		Case "$show_zt"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_zt(para(0),para(1),para(2),para(3),para(4),para(5),para(6),para(7),para(8)))
			If Err Then
				Response.Write "<br/>$show_zt$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If

		Case "$show_m"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_m(para(0),para(1),para(2),para(3),para(4),para(5),para(6),para(7),para(8)))
			If Err Then
				Response.Write "<br/>$show_m$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
			''''''''''''''''''''''''''''
		Case "$show_userlog"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_userlog(para(0),para(1),para(2),para(3),para(4),para(5)))
			If Err Then
				Response.Write Err.Description
				Response.Write "<br/>$show_userlog$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_comment"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_comment(para(0),para(1)))
			If Err Then
				Response.Write "<br/>$show_comment$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_subject"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_subject(para(0)))
			If Err Then
				Response.Write "<br/>$show_subject$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_blogupdate"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_blogupdate(para(0)))
			If Err Then
				Response.Write "<br/>$show_blogupdate$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_newblogger"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_newblogger(para(0)))
			If Err Then
				Response.Write "<br/>$show_newblogger$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_bestblog"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_bestblog(para(0)))
			If Err Then
				Response.Write "<br/>$show_bestblog$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_bloger"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_bloger(para(0)))
			If Err Then
				Response.Write "<br/>$show_bloger$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_class"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_class(para(0)))
			If Err Then
				Response.Write "<br/>$show_class$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_search"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_search(para(0)))
			If Err Then
				Response.Write "<br/>$show_search$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_cityblogger"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_cityblogger(para(0)))
			If Err Then
				Response.Write "<br/>$show_cityblogger$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_newphoto"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_newphoto(para(0),para(1),para(2),para(3)))
			If Err Then
				Response.Write Err.Description
				Response.Write "<br/>$show_newphoto$标签有错误，请检查参数"
				Response.End()
			End If
		Case "$show_blogstar2"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",show_blogstar2(para(0),para(1),para(2),para(3)))
			If Err Then
				Response.Write "<br/>$show_blogstar2$标签有错误，请检查参数"
				Response.End()
			End If
		 Case "$show_hottag"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetHotTags(para(0),para(1),para(2),para(3)))
		Case "$show_treeclass"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",TreeClass(para(0)))
			If Err Then
				Response.Write "<br/>$show_treeclass$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_teams"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetTeams(para(0),para(1),para(2),para(3),para(4),para(5)))
			If Err Then
				Response.Write "<br/>$show_teams$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_posts"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetPosts(para(0),para(1),para(2),para(3),para(4)))
			If Err Then
				Response.Write "<br/>$show_posts$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_hotblog"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetHotUsers(para(0),para(1)))
			If Err Then
				Response.Write "<br/>$show_hotblog$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_bl"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetArgueList(para(0),para(1),para(2)))
			If Err Then
				Response.Write "<br/>$show_bl$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_template"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetTemplate(para(0)))
			If Err Then
				Response.Write "<br/>$show_template$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_album"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetAlbum(para(0),para(1)))
			If Err Then
				Response.Write "<br/>$show_album$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_pic"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetPic(para(0),para(1)))
			If Err Then
				Response.Write "<br/>$show_pic$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_diggs"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetDiggs(para(0),para(1)))
			If Err Then
				Response.Write "<br/>$show_diggs$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_userdiggs"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetUserDiggs(para(0),para(1)))
			If Err Then
				Response.Write "<br/>$show_diggs$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		Case "$show_rnduser"
			G_P_Show=Replace(G_P_Show,label&"("&tmpstr&")$",GetRndUser(para(0),para(1),para(2),para(3),para(4),para(5)))
			If Err Then
				Response.Write "<br/>$show_rnduser$标签有错误，请检查参数"
				Response.Write Err.Description
				Response.End()
			End If
		End select
		tmp2 = 1
	Wend
End Sub
'-=====================================================================================


Function show_zt(n, l, order, action, sdate, classid, classname, subjectname, info)
	show_zt = ""
	Dim rs, msql, ordersql, actionsql, classsql, rstmp, i, postname, posttime, userurl, ustr
	Dim arrayList,arrayListTemp
	ReDim arrayList(n-1)
	i = 0
	select Case order
	Case 1
	   ordersql = " order by logid desc"
	   'ordersql = " order by addtime desc"
	Case 2
		ordersql = " order by iis desc,logid DESC"
	Case 3
		ordersql = " order by commentnum desc,logid DESC"
	End select

	select Case action
	Case 1
		actionsql = ""
	Case 2
		actionsql = " and isbest=1"
	End select
	If classid = 0 Then
		classsql = ""
	Else
		If InStr(classid,"|") Then
		Dim o
		For o=0 To UBound(Split(classid,"|",-1,1))

		If o=0 Then
			ustr = ustr&" and ("
		Else
			ustr = ustr&" or "
		End If
		ustr = ustr&" In_feature_list like '%,"&Split(classid,"|",-1,1)(o)&",%'"

		Next
		ustr=ustr&" ) "
				Else
		ustr=" and In_feature_list like '%,"&Split(classid,"|",-1,1)(o)&",%' "
		End If
		classsql=ustr
		'classsql=" and classid="&CLng(classid)
	End If
	msql="select top "&n&" topic,logfile,addtime,commentnum,iis,logid,classid,subjectid,author,userid from oblog_log where (IsSpecial = 0 OR IsSpecial IS NULL) And isdraft=0 and passcheck=1 And oblog_log.isdel=0 and (oblog_log.is_log_default_hidden=0 or oblog_log.is_log_default_hidden is null) "
	If Is_Sqldata = 0 Then
		msql = msql&" and datediff('d',oblog_log.truetime,Now())<"&Int(sdate)
	Else
		sdate = DateAdd("d",-1*Abs(sdate),Now())
		sdate = GetDateCode(sdate,0)
		msql = msql&" and truetime>'"&sdate&"'"
	End IF
	msql=msql&actionsql&classsql
	msql=msql&ordersql
'	OB_DEBUG msql, 1
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open msql ,CONN,1,1
	show_zt=show_zt & vbcrlf & "<ul>" & vbcrlf
	Do While Not rs.EOF
		arrayList(i) = rs("userid")
		show_zt=show_zt&"	<li>"
		posttime = rs(2)
		If classname = 1 Then
			Set rstmp = oblog.execute("select id,classname from oblog_logclass where id=" & rs(6))
			If Not rstmp.EOF Then
				show_zt=show_zt&"<a class=""oblog_class"" href=""list.asp?classid="&rstmp(0)&""" target=""_blank"">〖"&rstmp(1)&"〗</a>"
			End If
		End If
		If subjectname = 1 Then
			Set rstmp = oblog.execute("select subjectid,subjectname from oblog_subject where subjectid=" & rs(7))
			If Not rstmp.EOF Then
				show_zt=show_zt&"<a class=""oblog_subject"" href=""blog.asp?name="&rs("author")&"&subjectid="&rstmp(0)&""" target=""_blank"">["&oblog.filt_html(rstmp(1))&"]</a>"
			End If
		End If
		Dim topic
		If rs(0) <> "" Then
			topic = Replace(rs(0), "'", "")
			If topic <> "" Then
				If oblog.strLength(topic) > Int(l) Then
					topic = oblog.InterceptStr(topic, Int(l) - 3) & "..."
				End If
			End If
		End If
		show_zt=show_zt&"<a href="""&rs(1)&""" title="""&oblog.filt_html(rs(0))&""" target=_blank>"&oblog.filt_html(topic)&"</a>"
'		If oblog.cacheConfig(5) = 1 Then
'			userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
'		Else
'			userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
'		End If
		userurl = blogurl & "go.asp?userid="&rs("userid")
		postname = "<span name=""nickname_"&rs("userid")&""" id=""nickname_"&rs("userid")&""">"&rs("userid")&"</span>"
'		postname = "nickname_"&rs("userid")
		select Case CInt(info)
			Case 1
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_time"">"&formatdatetime(posttime,0)&"</span><span class=""ob_log_c1"">)</span>"
			Case 2
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_time"">"&posttime&"</span><span class=""ob_log_c1"">)</span>"
			Case 3
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c1"">)</span>"
			Case 4
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_num"">"&rs(4)&"</span><span class=""ob_log_c1"">)</span>"
			Case 5
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_num"">"&rs(4)&"</span><span class=""ob_log_c1"">)</span>"
			Case 6
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_time"">"&formatdatetime(posttime,1)&"</span><span class=""ob_log_c1"">)</span>"
			Case 7
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_time"">"&formatdatetime(posttime,1)&"</span><span class=""ob_log_c1"">)</span>"
			Case 8
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_num"">"&rs(3)&"</span><span class=""ob_log_c1"">)</span>"
			Case 9
				show_zt=show_zt&"<span class=""ob_log_c1"">(</span><span class=""ob_log_bname"">"&oblog.filt_html(rs("blogname"))&"<span><span class=""ob_log_c1"">)</span>"
			Case Else
		End select
		show_zt=show_zt&"</li>" & vbcrlf
		rs.MoveNext
		i = i + 1
		If i >= Int(n) Then Exit Do
	Loop
	show_zt=show_zt & "</ul>" & vbcrlf

	If Not rs.Bof Then
		'info参数为1，3，4，6之一才需调用用户昵称
		If InStr("1,3,4,6",info) > 0 Then
			show_zt = show_zt & oblog.GetNickNameById (arrayList,i,n&l&order&action&sdate&classid&classname&subjectname&info)
'			show_zt = oblog.GetNameNameByUserId(arrayList,show_zt)
		End if
	End If
	Set rs = Nothing
	Set rstmp = Nothing
End Function


Function show_m(n, l, order, action, sdate, classid, classname, subjectname, info)
	show_m = ""
	Dim rs, msql, ordersql, actionsql, classsql, rstmp, i, postname, posttime, userurl, ustr
	Dim arrayList,arrayListTemp
	ReDim arrayList(n-1)
	i = 0
	select Case order
	Case 1
	   ordersql = " order by logid desc"
	   'ordersql = " order by addtime desc"
	Case 2
		ordersql = " order by iis desc,logid DESC"
	Case 3
		ordersql = " order by commentnum desc,logid DESC"
	End select

	select Case action
	Case 1
		actionsql = ""
	Case 2
		actionsql = " and isbest=1"
	End select
	If classid = 0 Then
		classsql = ""
	Else
		If InStr(classid,"|") Then
		Dim o
		For o=0 To UBound(Split(classid,"|",-1,1))

		If o=0 Then
			ustr = ustr&" and ("
		Else
			ustr = ustr&" or "
		End If
		ustr = ustr&" Magazine_list like '%,"&Split(classid,"|",-1,1)(o)&",%' "

		Next
		ustr=ustr&" ) "
		Else
		ustr=" and Magazine_list like '%,"&Split(classid,"|",-1,1)(o)&",%' "
		End If
		classsql=ustr
		'classsql=" and classid="&CLng(classid)
	End If
	msql="select top "&n&" topic,logfile,addtime,commentnum,iis,logid,classid,subjectid,author,userid from oblog_log where (IsSpecial = 0 OR IsSpecial IS NULL) And isdraft=0 and passcheck=1 And oblog_log.isdel=0 and (oblog_log.is_log_default_hidden=0 or oblog_log.is_log_default_hidden is null) "
	If Is_Sqldata = 0 Then
		msql = msql&" and datediff('d',oblog_log.truetime,Now())<"&Int(sdate)
	Else
		sdate = DateAdd("d",-1*Abs(sdate),Now())
		sdate = GetDateCode(sdate,0)
		msql = msql&" and truetime>'"&sdate&"'"
	End IF
	msql=msql&actionsql&classsql
	msql=msql&ordersql
'	OB_DEBUG msql, 1
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open msql ,CONN,1,1
	show_m=show_m & vbcrlf & "<ul>" & vbcrlf
	Do While Not rs.EOF
		arrayList(i) = rs("userid")
		show_m=show_m&"	<li>"
		posttime = rs(2)
		If classname = 1 Then
			Set rstmp = oblog.execute("select id,classname from oblog_logclass where id=" & rs(6))
			If Not rstmp.EOF Then
				show_m=show_m&"<a class=""oblog_class"" href=""list.asp?classid="&rstmp(0)&""" target=""_blank"">〖"&rstmp(1)&"〗</a>"
			End If
		End If
		If subjectname = 1 Then
			Set rstmp = oblog.execute("select subjectid,subjectname from oblog_subject where subjectid=" & rs(7))
			If Not rstmp.EOF Then
				show_m=show_m&"<a class=""oblog_subject"" href=""blog.asp?name="&rs("author")&"&subjectid="&rstmp(0)&""" target=""_blank"">["&oblog.filt_html(rstmp(1))&"]</a>"
			End If
		End If
		Dim topic
		If rs(0) <> "" Then
			topic = Replace(rs(0), "'", "")
			If topic <> "" Then
				If oblog.strLength(topic) > Int(l) Then
					topic = oblog.InterceptStr(topic, Int(l) - 3) & "..."
				End If
			End If
		End If
		show_m=show_m&"<a href="""&rs(1)&""" title="""&oblog.filt_html(rs(0))&""" target=_blank>"&oblog.filt_html(topic)&"</a>"
'		If oblog.cacheConfig(5) = 1 Then
'			userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
'		Else
'			userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
'		End If
		userurl = blogurl & "go.asp?userid="&rs("userid")
		postname = "<span name=""nickname_"&rs("userid")&""" id=""nickname_"&rs("userid")&""">"&rs("userid")&"</span>"
'		postname = "nickname_"&rs("userid")
		select Case CInt(info)
			Case 1
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_time"">"&formatdatetime(posttime,0)&"</span><span class=""ob_log_c1"">)</span>"
			Case 2
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_time"">"&posttime&"</span><span class=""ob_log_c1"">)</span>"
			Case 3
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c1"">)</span>"
			Case 4
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_num"">"&rs(4)&"</span><span class=""ob_log_c1"">)</span>"
			Case 5
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_num"">"&rs(4)&"</span><span class=""ob_log_c1"">)</span>"
			Case 6
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_time"">"&formatdatetime(posttime,1)&"</span><span class=""ob_log_c1"">)</span>"
			Case 7
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_time"">"&formatdatetime(posttime,1)&"</span><span class=""ob_log_c1"">)</span>"
			Case 8
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_num"">"&rs(3)&"</span><span class=""ob_log_c1"">)</span>"
			Case 9
				show_m=show_m&"<span class=""ob_log_c1"">(</span><span class=""ob_log_bname"">"&oblog.filt_html(rs("blogname"))&"<span><span class=""ob_log_c1"">)</span>"
			Case Else
		End select
		show_m=show_m&"</li>" & vbcrlf
		rs.MoveNext
		i = i + 1
		If i >= Int(n) Then Exit Do
	Loop
	show_m=show_m & "</ul>" & vbcrlf

	If Not rs.Bof Then
		'info参数为1，3，4，6之一才需调用用户昵称
		If InStr("1,3,4,6",info) > 0 Then
			show_m = show_m & oblog.GetNickNameById (arrayList,i,n&l&order&action&sdate&classid&classname&subjectname&info)
'			show_m = oblog.GetNameNameByUserId(arrayList,show_m)
		End if
	End If
	Set rs = Nothing
	Set rstmp = Nothing
End Function



'-=====================================================================================
Function show_log(n, l, order, action, sdate, classid, classname, subjectname, info)
	show_log = ""
	Dim rs, msql, ordersql, actionsql, classsql, rstmp, i, postname, posttime, userurl, ustr
	Dim arrayList,arrayListTemp
	ReDim arrayList(n-1)
	i = 0
	select Case order
	Case 1
	   ordersql = " order by logid desc"
	   'ordersql = " order by addtime desc"
	Case 2
		ordersql = " order by iis desc,logid DESC"
	Case 3
		ordersql = " order by commentnum desc,logid DESC"
	End select

	select Case action
	Case 1
		actionsql = ""
	Case 2
		actionsql = " and isbest=1"
	End select
	If classid = 0 Then
		classsql = ""
	Else
		set rs=oblog.execute("select id from oblog_logclass where parentpath like '"&classid&",%' OR parentpath like '%,"&classid&"' OR parentpath like '%,"&classid&",%'")
		While Not rs.EOF
			ustr=ustr&","&rs(0)
			rs.MoveNext
		Wend
		ustr=classid&ustr
		classsql=" and classid in ("&ustr&")"
		'classsql=" and classid="&CLng(classid)
	End If
	msql="select top "&n&" topic,logfile,addtime,commentnum,iis,logid,classid,subjectid,author,userid from oblog_log where (IsSpecial = 0 OR IsSpecial IS NULL) And isdraft=0 and passcheck=1 And oblog_log.isdel=0 and (oblog_log.is_log_default_hidden=0 or oblog_log.is_log_default_hidden is null) "
	If Is_Sqldata = 0 Then
		msql = msql&" and datediff('d',oblog_log.truetime,Now())<"&Int(sdate)
	Else
		sdate = DateAdd("d",-1*Abs(sdate),Now())
		sdate = GetDateCode(sdate,0)
		msql = msql&" and truetime>'"&sdate&"'"
	End IF
	msql=msql&actionsql&classsql
	msql=msql&ordersql
'	OB_DEBUG msql, 1
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open msql ,CONN,1,1
	show_log=show_log & vbcrlf & "<ul>" & vbcrlf
	Do While Not rs.EOF
		arrayList(i) = rs("userid")
		show_log=show_log&"	<li>"
		posttime = rs(2)
		If classname = 1 Then
			Set rstmp = oblog.execute("select id,classname from oblog_logclass where id=" & rs(6))
			If Not rstmp.EOF Then
				show_log=show_log&"<a class=""oblog_class"" href=""list.asp?classid="&rstmp(0)&""" target=""_blank"">〖"&rstmp(1)&"〗</a>"
			End If
		End If
		If subjectname = 1 Then
			Set rstmp = oblog.execute("select subjectid,subjectname from oblog_subject where subjectid=" & rs(7))
			If Not rstmp.EOF Then
				show_log=show_log&"<a class=""oblog_subject"" href=""blog.asp?name="&rs("author")&"&subjectid="&rstmp(0)&""" target=""_blank"">["&oblog.filt_html(rstmp(1))&"]</a>"
			End If
		End If
		Dim topic
		If rs(0) <> "" Then
			topic = Replace(rs(0), "'", "")
			If topic <> "" Then
				If oblog.strLength(topic) > Int(l) Then
					topic = oblog.InterceptStr(topic, Int(l) - 3) & "..."
				End If
			End If
		End If
		show_log=show_log&"<a href="""&rs(1)&""" title="""&oblog.filt_html(rs(0))&""" target=_blank>"&oblog.filt_html(topic)&"</a>"
'		If oblog.cacheConfig(5) = 1 Then
'			userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
'		Else
'			userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
'		End If
		userurl = blogurl & "go.asp?userid="&rs("userid")
		postname = "<span name=""nickname_"&rs("userid")&""" id=""nickname_"&rs("userid")&""">"&rs("userid")&"</span>"
'		postname = "nickname_"&rs("userid")
		select Case CInt(info)
			Case 1
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_time"">"&formatdatetime(posttime,0)&"</span><span class=""ob_log_c1"">)</span>"
			Case 2
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_time"">"&posttime&"</span><span class=""ob_log_c1"">)</span>"
			Case 3
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c1"">)</span>"
			Case 4
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_num"">"&rs(4)&"</span><span class=""ob_log_c1"">)</span>"
			Case 5
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_num"">"&rs(4)&"</span><span class=""ob_log_c1"">)</span>"
			Case 6
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_uname""><a href="&userurl&" target=_blank>"&postname&"</a></span><span class=""ob_log_c2"">,</span><span class=""ob_log_time"">"&formatdatetime(posttime,1)&"</span><span class=""ob_log_c1"">)</span>"
			Case 7
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_time"">"&formatdatetime(posttime,1)&"</span><span class=""ob_log_c1"">)</span>"
			Case 8
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_num"">"&rs(3)&"</span><span class=""ob_log_c1"">)</span>"
			Case 9
				show_log=show_log&"<span class=""ob_log_c1"">(</span><span class=""ob_log_bname"">"&oblog.filt_html(rs("blogname"))&"<span><span class=""ob_log_c1"">)</span>"
			Case Else
		End select
		show_log=show_log&"</li>" & vbcrlf
		rs.MoveNext
		i = i + 1
		If i >= Int(n) Then Exit Do
	Loop
	show_log=show_log & "</ul>" & vbcrlf

	If Not rs.Bof Then
		'info参数为1，3，4，6之一才需调用用户昵称
		If InStr("1,3,4,6",info) > 0 Then
			show_log = show_log & oblog.GetNickNameById (arrayList,i,n&l&order&action&sdate&classid&classname&subjectname&info)
'			show_log = oblog.GetNameNameByUserId(arrayList,show_log)
		End if
	End If
	Set rs = Nothing
	Set rstmp = Nothing
End Function

Function show_userlog(userid,n, l, order,  subjectid,  info)
	Dim rs, strSql, strOrderSql, i, posttime,topic,strContent
	i = 0
	select Case order
		Case 1
			'strOrderSql = " order by logid desc"
			strOrderSql = " order by addtime desc"
		Case 2
			strOrderSql = " order by iis desc,logid DESC"
		Case 3
			strOrderSql = " order by commentnum desc,logid DESC"
	End select

	strSql = "select Top "&n&" topic,logfile,addtime,commentnum,iis,logid,author,userid"
	strSql = strSql & " from oblog_log where (IsSpecial = 0 OR IsSpecial IS NULL) And isdraft=0 and passcheck=1 and isdel=0 "
	'过滤掉当日之后的日志
	strSql = strSql & " And addtime< "&G_Sql_Now
	If subjectid<>0 And IsNumeric(subjectid) Then
		strSql = strSql & " And Subjectid=" & CLng(subjectid)
	End If
	strSql = strSql & " And UserId=" & CLng(Userid) & strOrderSql
	'Response.Write strSql
	Set rs = oblog.execute(strSql)
	strContent= vbcrlf & "<ul>" & vbcrlf
	Do While Not rs.EOF
		strContent=strContent&"	<li>"
		posttime = rs("addtime")
		If rs("topic") <> "" Then
			topic = Replace(rs("topic"), "'", "")
			If topic <> "" Then
				If oblog.strLength(topic) > Int(l) Then
					topic = oblog.InterceptStr(topic, Int(l) - 3) & "..."
				End If
			End If
		End If
		strContent=strContent&"<a href=""go.asp?logid=" & rs("logid") &""" title="""&oblog.filt_html(rs(0))&""" target=_blank>"&oblog.filt_html(topic)&"</a>"
		If  CInt(info)=1 Then
			strContent=strContent&"("&formatdatetime(posttime,1)&")"
		End If
		strContent=strContent&"</li>" & vbcrlf
		rs.MoveNext
		i = i + 1
		If i >= Int(n) Then Exit Do
	Loop
	show_userlog=strContent & "</ul>" & vbcrlf
	strContent=""
	Set rs = Nothing
End Function

Function show_class(m)
	Dim rs
	'show_class="<a href=index.asp>首页("&blogcount&")</a><br>"
	Dim i, brstr
	show_class = ""
	m = Int(m)
	Set rs = oblog.execute("select id,classname from oblog_logclass Where idtype=0 And child=0 and parentid=0 order by RootID,OrderID")
	If m = 0 Then
		While Not rs.EOF
			show_class=show_class&"<a href=""list.asp?classid="&rs(0)&""" title="""&rs(1)&""">"&rs(1)&"</a><br />" & vbcrlf
			rs.MoveNext
		Wend
	Else
		i = 0
		While Not rs.EOF
			i = i + 1
			If i = Int(m) Then
				brstr = "<br />" & vbcrlf
				i = 0
			Else
				brstr = ""
			End If
			show_class=show_class&"<a href=""list.asp?classid="&rs(0)&""" title="""&rs(1)&""">"&rs(1)&"</a>&nbsp;" & brstr & vbcrlf
			rs.MoveNext
		Wend
		if right(show_class,6)="<br />" then show_class=left(show_class,len(show_class)-6)
	End If
	Set rs = Nothing
End Function

Function show_comment(n, l)
	Dim rs
	set rs=oblog.execute("select top "&n&" mainid,commenttopic,comment_user,addtime,commentid from [oblog_comment] where isdel=0 order by commentid desc")
	show_comment= vbcrlf & "<ul>" & vbcrlf
	While Not rs.EOF
		show_comment=show_comment&"	<li><a href=""go.asp?logid="&rs(0)&"&commentid="&rs(4)&""" target=""_blank"" title="""&oblog.filt_html(rs(2))&"回复于"&rs(3)&""">"&oblog.InterceptStr(oblog.filt_html(rs(1)),CLng(l))&"</a></li>" & vbcrlf
		rs.MoveNext
	Wend
	show_comment=show_comment&"</ul>" & vbcrlf
	Set rs = Nothing
End Function

Function show_subject(n)
	Dim i, rs
	i = 0
	'set rs=oblog.execute("select top "&n&" subjectid,oblog_subject.userid,subjectname,subjectlognum,user_dir,user_folder from [oblog_subject],oblog_user where oblog_subject.userid=oblog_user.userid and oblog_subject.oblog_subjecttype=0 order by subjectlognum desc")
	set rs=oblog.execute("select a.*,b.username From (select top " & n &" Subjectid,SubjectName,SubjectlogNum,userid From oBlog_subject where subjecttype=0 order by subjectlognum desc) a ,oblog_user b Where a.userid=b.userid")
	show_subject= vbcrlf & "<ul>" & vbcrlf
	Do While Not rs.EOF
		show_subject=show_subject&"	<li><a href=""blog.asp?name="&rs("username")&"&subjectid="&rs("subjectid")&""" target=""_blank"" title="""&oblog.filt_html(rs("subjectname"))&"("&rs("SubjectlogNum")&")"">"&oblog.filt_html(rs("subjectname"))&"("&rs("SubjectlogNum")&")</a></li>" & vbcrlf
		rs.MoveNext
		i = i + 1
		If i >= Int(n) Then Exit Do
	Loop
	show_subject=show_subject&"</ul>" & vbcrlf
	Set rs = Nothing
End Function

Function show_blogupdate(n)
	Dim i, rs, userurl
	i = 0
	set rs=oblog.execute("select top "&n&" username,log_count,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder from [oblog_user] where lockuser=0 and isdel=0 AND user_level >=7 order by log_count desc,userid DESC")
	show_blogupdate= vbcrlf & "<ul>" & vbcrlf
	Do While Not rs.EOF
		If oblog.cacheConfig(5) = 1 Then
			userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
		Else
			userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
		End If
		If rs("blogname") <> "" Then
			show_blogupdate=show_blogupdate&"	<li><a href="""&userurl&""" target=""_blank"" title="""&rs("blogname")&"("&rs("log_count")&")"">"&rs("blogname")&"("&rs("log_count")&")</a></li>" & vbcrlf
		Else
			show_blogupdate=show_blogupdate&"	<li><a href="""&userurl&""" target=""_blank"" title="""&rs("username")&"("&rs("log_count")&")"">"&rs("username")&"("&rs("log_count")&")</a></li>" & vbcrlf
		End If
		rs.MoveNext
		i = i + 1
		If i >= Int(n) Then Exit Do
	Loop
	show_blogupdate=show_blogupdate&"</ul>" & vbcrlf
	Set rs = Nothing
End Function

Function show_newblogger(n)
	Dim rs, userurl,userico
	set rs=oblog.execute("select top "&n&" username,log_count,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder from [oblog_user] where lockuser=0 and isdel=0 AND user_level >=7 order by userid desc")
	show_newblogger= vbcrlf & "<ul>" & vbcrlf
	While Not rs.EOF
		If oblog.cacheConfig(5) = 1 Then
			userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
		Else
			userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
		End If
		If rs(3) <> "" Then
			show_newblogger=show_newblogger&"	<li><a href="""&userurl&""" target=""_blank"" title="""&rs(3)&"("&rs(1)&")"">"&rs(3)&"("&rs(1)&")</a></li>" & vbcrlf
		Else
			show_newblogger=show_newblogger&"	<li><a href="""&userurl&""" target=""_blank"" title="""&rs(0)&"("&rs(1)&")"">"&rs(0)&"("&rs(1)&")</a></li>" & vbcrlf
		End If
		rs.MoveNext
	Wend
	show_newblogger=show_newblogger&"</ul>" & vbcrlf
	Set rs = Nothing
End Function

Function show_bestblog(n)
	Dim i, rs, userurl
	i = 0
	set rs=oblog.execute("select top "&n&" username,log_count,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder from [oblog_user] where user_isbest=1 and isdel=0 AND user_level >=7 order by log_count desc,userid DESC")
	show_bestblog= vbcrlf & "<ul>" & vbcrlf
	Do While Not rs.EOF
		If oblog.cacheConfig(5) = 1 Then
			userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
		Else
			userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
		End If
		If rs(3) <> "" Then
			show_bestblog=show_bestblog&"	<li><a href="""&userurl&""" target=""_blank"" title="""&rs(3)&"("&rs(1)&")"">"&rs(3)&"("&rs(1)&")</a></li>" & vbcrlf
		Else
			show_bestblog=show_bestblog&"	<li><a href="""&userurl&""" target=""_blank"" title="""&rs(0)&"("&rs(1)&")"">"&rs(0)&"("&rs(1)&")</a></li>" & vbcrlf
		End If
		rs.MoveNext
		i = i + 1
		If i >= Int(n) Then Exit Do
	Loop
	show_bestblog=show_bestblog&"</ul>" & vbcrlf
	Set rs = Nothing
End Function

Function show_count()
	Dim rs,logToday
	If Is_Sqldata = 0 Then
		Set rs = oblog.execute("select COUNT(logid) FROM oblog_log WHERE DATEDIFF('d',truetime,Now)=0 AND isdel=0 ")
	Else
		Set rs = oblog.execute("select COUNT(logid) FROM oblog_log WHERE truetime>=CONVERT(CHAR(10),GETDATE(),120) AND truetime < CONVERT(CHAR(10),GETDATE()+1,120) AND isdel=0 ")
	End if
	logToday=rs(0)
	Set rs = Nothing
	show_count = vbcrlf & "<ul id=""blog_info"">" & vbcrlf
	show_count = show_count & "	<li>博客：" & oblog.setup(4, 0) & "</li>" & vbcrlf
	show_count = show_count & "	<li>日志："&oblog.setup(1,0)&"</li>" & vbcrlf
	show_count = show_count & "	<li>评论："&oblog.setup(2,0)&"</li>" & vbcrlf
	show_count = show_count & "	<li>留言："&oblog.setup(3,0)&"</li>" & vbcrlf
	show_count = show_count & "	<li>昨日："&OB_IIF(oblog.setup(10, 0),0)&"</li>" & vbcrlf
	show_count = show_count & "	<li>今日："&logToday&"</li></ul>" & vbcrlf
	show_count = show_count & "</ul>" & vbcrlf
End Function

Function show_sysxml()
	show_sysxml = "<a href=""" & blogdir & "rssfeed.asp"" target=""_blank"" title=""订阅本站最新文章""><img src=""Images/xml.gif"" width=""36"" height=""14"" border=""0""></a>"
End Function

Function show_friends()
	show_friends=OB_IIF(oblog.setup(6, 0),"")
End Function

Function show_placard()
	show_placard=OB_IIF(oblog.setup(5, 0),"")
End Function

Function show_bloger(m)
	Dim rs
	Dim i, brstr
	m = Int(m)
	Set rs = oblog.execute("select id,classname from oblog_userclass order by RootID,OrderID")
	If m = 0 Then
		While Not rs.EOF
			show_bloger=show_bloger&"<a href=""listblogger.asp?usertype="&rs(0)&""" title="""&rs(1)&""">"&rs(1)&"</a><br/>" & vbcrlf
			rs.MoveNext
		Wend
	Else
		i = 0
		While Not rs.EOF
			i = i + 1
			If i = Int(m) Then
				brstr = "<br/>"
				i = 0
			Else
				brstr = ""
			End If
			show_bloger=show_bloger&"<a href=listblogger.asp?usertype="&rs(0)&" title="""&rs(1)&""">"&rs(1)&"</a>&nbsp;" & brstr & vbcrlf
			rs.MoveNext
		Wend
	End If
	Set rs = Nothing
End Function

Function show_search(i)
	If i = 0 Then i = "" Else i = "<br />"
	show_search = vbcrlf & "<form name=""search"" method=""post"" action=""list.asp"">" & vbcrlf
	show_search = show_search & "	<select name=""selecttype"" id=""selecttype"">" & vbcrlf
	show_search = show_search & "		<option value=""topic"" selected>日志标题</option>" & vbcrlf
	show_search = show_search & "		<option value=""logtext"">日志内容</option>" & vbcrlf
	show_search = show_search & "		<option value=""id"">博客名称</option>" & i & vbcrlf
	show_search = show_search & "		<option value=""username"">用户名</option>" & i & vbcrlf
	show_search = show_search & "		<option value=""nickname"">用户昵称</option>" & i & vbcrlf
	show_search = show_search & "	</select>" & i & vbcrlf
	show_search = show_search & "	<input name=""keyword"" type=""text"" id=""keyword"" size=""16"" maxlength=""40"">" & vbcrlf
	show_search = show_search & "	<input type=""submit"" name=""Submit"" id=""Submit"" value=""搜索"">" & vbcrlf
	show_search = show_search & "</form>" & vbcrlf
End Function

Function show_cityblogger(i)
	show_cityblogger = vbcrlf & "<form name=""oblogform"" id=""cityblogger"" action=""listblogger.asp"">" & vbcrlf
	show_cityblogger = show_cityblogger & oblog.type_city("", "") & vbcrlf
	show_cityblogger = show_cityblogger &"	<input type=""submit"" value=""搜索"">" & vbcrlf
	show_cityblogger = show_cityblogger &"</form>" & vbcrlf

	If i = 1 Then show_cityblogger = Replace(show_cityblogger, "<select name=""city""", "<br /><select name=""city""")
End Function

Function show_newphoto(n, i, w, h)
	Dim rs, sReadMe,imgsrc,fso,wstr,hstr,j,preImgSrc
	Set fso = Server.CreateObject(oblog.CacheCompont(1))
'	If i = 1 Then i = "<br />" Else i = ""
	'兼容4.0模版
	If i = 0 Then i = 4
	if w<>0 or w<>"" then wstr="width="""&w&""""
	if h<>0 or h<>"" then hstr="height="""&h&""""
	Set rs = oblog.execute("select TOP "&N&" c.photo_path,c.photo_readme,c.userid,c.fileid FROM oblog_album c where (c.ishide = 0 OR c.ishide IS NULL) order by photoid desc")
	While Not rs.EOF
		j = j + 1
		If IsNull(rs(1)) Then
			sReadMe = ""
		Else
			sReadMe = oblog.filt_html(rs(1))
		End If
		imgsrc=rs(0)

		preImgSrc=Replace(imgsrc,right(imgsrc,3),"jpg")
		preImgSrc=Replace(preImgSrc,right(preImgSrc,len(preImgSrc)-InstrRev(preImgSrc,"/")),"pre"&right(preImgSrc,len(preImgSrc)-InstrRev(preImgSrc,"/")))
		if  not fso.FileExists(Server.MapPath(preImgSrc)) then
			preImgSrc=imgsrc
		end if
		show_newphoto=show_newphoto&"<a href=""go.asp?fileid="&rs("fileid")&""" target=""_blank""><img src="""&preImgSrc&""" "&wstr&" "&hstr&" hspace=""6"" border=""0"" vspace=""6"" alt="""& sReadMe &""" /></a>"
		If j Mod i = 0 Then show_newphoto=show_newphoto& "<br />"
		rs.MoveNext
	Wend
	Set rs = Nothing
End Function

Function show_blogstar()
	Dim rs
	Set rs = oblog.execute("select top 1 * from oblog_blogstar where ispass=1 order by id desc")
	If Not rs.EOF Then
		show_blogstar = vbcrlf & "<div id=""blogstar"">" & vbcrlf
		show_blogstar = show_blogstar & "	<div class=""blogstarimg""><a href=""" & rs("userurl") & """ target=""_blank""><img src=""" & rs("picurl") & """  hspace=""3"" border=""0"" vspace=""3"" alt=""" & oblog.filt_html(rs("blogname")) & """ /></a></div>" & vbcrlf
		show_blogstar=show_blogstar & "	<div class=""blogstarname"">博客："&"<a href="""&rs("userurl")&""" target=""_blank"" title="""&oblog.filt_html(rs("blogname"))&""">"&oblog.filt_html(rs("blogname"))&"</a></div>" & vbcrlf
		show_blogstar = show_blogstar & "	<div class=""blogstarinfo"">简介："&oblog.filt_html(rs("info"))&"</div>" & vbcrlf
		show_blogstar = show_blogstar & "</div>" & vbcrlf
	Else
		show_blogstar = " "
	End If
	Set rs = Nothing
End Function

Public Function show_blogstar2(iNumber, iPerline, iWidth, iHeight)
	Dim rs, iCount, sLine
	If Not IsNumeric(iNumber) Then
		iNumber = 1
	Else
		iNumber = CLng(iNumber)
	End If
	'iWidth=160
	'iHeight=160
	If iNumber = 0 Then iNumber = 1
	Set rs = oblog.execute("select top " & iNumber & " * from oblog_blogstar where ispass=1 order by id desc")
	If Not rs.EOF Then
		show_blogstar2 = vbcrlf & "<table id=""blogstar"" style=""table-layout: fixed; word-break: break-all; "" width=""100%"" border=""0"">" & vbcrlf
		show_blogstar2 = show_blogstar2 & "	<tr>" & vbcrlf


		If iNumber = 1 Then
			sLine = "		<td valign=""top"">" & vbcrlf
			sLine = sLine & "			<div class=""blogstarimg""><a href=""" & rs("userurl") & """ target=""_blank""><img src=""" & rs("picurl") & """ hspace=""3"" border=""0"" vspace=""3"" alt=""" & Left(oblog.filt_html(rs("blogname")) ,999)& """ onload=""javascript:if(this.width>" & iWidth & ") this.style.width=" & iWidth & ";"" /></a></div>" & vbcrlf
			sLine = sLine & "			<div class=""blogstarname"">博客：" & "<a href=""" & rs("userurl") & """ target=""_blank"" title=""" & Left(oblog.filt_html(rs("blogname")) ,999) & """>" & Left(oblog.filt_html(rs("blogname")) ,50) & "</a></div>" & vbcrlf
			sLine = sLine & "			<div class=""blogstarinfo"">简介：" & oblog.filt_html(rs("info")) & "</div>" & vbcrlf
			sLine = sLine & "		</td>" & vbcrlf
			show_blogstar2 = show_blogstar2 & sLine & "	</tr>" & vbCrLf
		'多图片时强制大小统一
		Else
			iCount = 1
			Do While Not rs.EOF
				sLine = "		<td valign=""top"">" & vbcrlf
				sLine = sLine & "			<div class=""blogstarimg""><a href=""" & rs("userurl") & """ target=""_blank""><img src=""" & rs("picurl") & """ hspace=""3"" border=""0"" vspace=""3"" alt=""" & Left (oblog.filt_html(rs("blogname")),999 )& """ width=" & iWidth & " height=" & iHeight & " /></a></div>" & vbcrlf
				sLine = sLine & "			<div class=""blogstarname"">博客：" & "<a href=""" & rs("userurl") & """ target=""_blank"" title=""" & Left (oblog.filt_html(rs("blogname")) ,999)& """>" & Left (oblog.filt_html(rs("blogname")) ,50)& "</a></div>" & vbcrlf
				sLine = sLine & "			<div class=""blogstarinfo"">简介：" & oblog.filt_html(rs("info")) & "</div>" & vbcrlf
				sLine = sLine & "		</td>" & vbCrLf
				show_blogstar2 = show_blogstar2 & sLine
				If iCount Mod iPerline = 0 Then show_blogstar2 = show_blogstar2 & "	</tr>" & vbcrlf
				iCount = iCount + 1
				rs.MoveNext
			Loop
			If Right(show_blogstar2, 5) <> "	</tr>" Then show_blogstar2 = show_blogstar2 & "	</tr>" & vbcrlf
		End If
		show_blogstar2 = show_blogstar2 & "</table>" & vbcrlf
	Else
		show_blogstar2 = " "
	End If
	rs.Close
	Set rs = Nothing
End Function

'获取标签
's 表现形式 1-列表形式,2-云图形式
'n 标签数目
'x 取消（防止用户启用此标签改动麻烦，函数暂不变）
'y 每行显示数目
Function GetHotTags(s,n,x,y)
	Dim sContent,sSql,rst,iFont,iFontSize,i,iFontFamily
	Dim sSplit
	sSplit="&nbsp;&nbsp;" & vbcrlf
	sSql="select * From (SELECT TOP "& n & " * FROM Oblog_Tags ORDER BY iNum DESC,tagid DESC) AS T Where iNum>0 AND iState=1 "
	If s=1 Then
		sSql= sSql & " Order By iNum Desc,tagid DESC "
	Else
		If Is_Sqldata = 1 Then
			sSql= sSql & " Order By Newid()"
		Else
			Randomize
			sSql= sSql & " Order By Rnd(-(TagID+"&Rnd()&"))"
		End If
	End if
	Set rst=conn.Execute(sSql)
	If rst.Eof Then
		sContent=""
	Else
		Do While Not rst.Eof
			If s=1 Then
				sContent= sContent & "<span><a href=""tags.asp?tagid=" & rst("tagID") &""">" & rst("Name")& "<span>(" & rst("iNum") &  ")</span></a></span>" & sSPlit
			Else
				Dim className,FontSize,FontWeight
				iFont=rst("iNum")
				If iFont <= 10 Then
					className = """tag_font1"""
					FontSize = "12px"
					FontWeight = "400"
				ElseIf iFont <=30 Then
					className = """tag_font2"""
					FontSize = "14px"
					FontWeight = "400"
				ElseIf iFont <=100 Then
					className = """tag_font3"""
					FontSize = "15px"
					FontWeight = "400"
				ElseIf iFont <=200 Then
					className = """tag_font4"""
					FontSize = "16px"
					FontWeight = "600"
				ElseIf iFont <=300 Then
					className = """tag_font5"""
					FontSize = "17px"
					FontWeight = "600"
				ElseIf iFont <=450 Then
					className = """tag_font6"""
					FontSize = "18px"
					FontWeight = "600"
				ElseIf iFont <=600 Then
					className = """tag_font7"""
					FontSize = "19px"
					FontWeight = "600"
				ElseIf iFont <=800 Then
					className = """tag_font8"""
					FontSize = "20px"
					FontWeight = "600"
				ElseIf iFont <=1000 Then
					className = """tag_font9"""
					FontSize = "21px"
					FontWeight = "600"
				Else
					className = """tag_font10"""
					FontSize = "22px"
					FontWeight = "600"
				End if
				if iFontSize >800 then iFontFamily="黑体"
				sContent= sContent & "<a href=""tags.asp?tagid=" & rst("tagID") & """ title=""TAG：" & Left(rst("Name"),10)& vbcrlf & "被使用"& rst("iNum") &"次""><span class="&className&" style=""font-size: "&FontSize&"; font-weight: "&FontWeight&"; font-family:"&iFontFamily&";"" >" & Left(rst("Name"),10)& "</span></a>" & sSPlit
			End If
			i=i+1
			If i Mod y = 0 Then
				sContent = sContent &  "<br />" & vbCrLf
			End If
			rst.Movenext
		Loop
	End If
	rst.Close
	Set rst=Nothing
	GetHotTags= sContent
	sContent=""
End Function

'x:1- 最新创建/2-最活跃群组(贴数最多)/3-规模大(人数最多) / 4-推荐群组
'n: 数目
'l: 题目显示长度
'y: 是否显示图标
'w:	图标宽度，不写则默认50
'h: 图标高度，不写则默认50
Function GetTeams(x,n,l,y,w,h)
	Dim rs,Sql,sRet,sIco
	Sql="select top " & n & " teamid,t_name,t_ico,icount0,(icount1+icount2) From oblog_team Where istate=3 and isdel=0  "
	select Case x
		Case 1
			Sql= Sql & " Order By teamid Desc"
		Case 2
			Sql= Sql & " Order By (icount1+icount2) Desc,teamid DESC"
		Case 3
			Sql= Sql & " Order By icount0 Desc,teamid DESC"
		Case 4
			Sql= Sql & " and isbest=1"
	End select
	Set rs=oblog.Execute(Sql)
	sRet= Vbcrlf & "<ul>" & Vbcrlf
	Do While Not rs.Eof
		sRet=sRet & "	<li>" & Vbcrlf
		If y=1 Then
			If w="" Then w=50:h=50
			sIco=LCase(Ob_IIF(rs(2),"images/ico_default.gif"))
			If Left(sico,7)<>"http://" Then sico=blogdir & sico
			sRet=sRet & "		<div class=""group_img""><a href=""group.asp?gid=" & rs(0) & """ target=""_blank""><img src=""" & sico & """ width=""" & w &""" height=""" & h &""" alt=""" & Left(oblog.filt_html((rs(1))),l) & "(" & rs(3) & "/" & rs(4) & ")"" /></a></div>" & Vbcrlf
		End if
		sRet=sRet & "		<div class=""group_name""><a href=""group.asp?gid=" & rs(0) & """ target=""_blank"" title=""" & Left(oblog.filt_html((rs(1))),l) & "(" & rs(3) & "/" & rs(4) & ")"">" & Left(oblog.filt_html((rs(1))),l) & "</a><span>(" & rs(3) & "/" & rs(4) & ")</span></div>" & Vbcrlf
		sRet=sRet & "	</li>" & Vbcrlf
		rs.movenext
	Loop
	Set rs=Nothing
	sRet=sRet & "</ul>" & Vbcrlf
	GetTeams=sRet
End Function

'获取群组文章
'teamid: 0 所有群组;如果是选择多个群组,则把群组ID用|分隔开,如1|2|8
'postnum: 帖子数目
'l:帖子主题显示字数
'u:是否显示用户名 0/1
't:是否显示发帖时间 0/1
Function GetPosts(teamid,postnum,l,u,t)
	Dim rs,sql,sRet,sAddon
	Dim arrayList,i
	ReDim arrayList(postnum-1)
	Sql="select Top " & postnum & " teamid,postid,topic,addtime,author,userid From oblog_teampost Where idepth=0 and isdel=0 "
	If teamid<>"" And teamid<>"0" Then
		teamid=Replace(teamid,"|",",")
		teamid  = FilterIDs(teamid)
		If teamid = "" Then Exit Function
		Sql=Sql & " And teamid In (" & teamid & ") "
	End If
	Sql=Sql & " Order by postid Desc"
	Set rs=oblog.Execute(Sql)
	sRet= Vbcrlf & "<ul>" & Vbcrlf
	i = 0
	If Not RS.Eof Then
		Do While Not rs.Eof
			arrayList(i) = rs("userid")
			sAddon=""
			sRet=sRet & "	<li><a href=""group.asp?gid=" & rs(0) & "&pid=" & rs(1) & """ target=""_blank"" title=""" & oblog.Filt_html(Left(OB_IIF(rs(2),"无题"),l)) & """>" & oblog.Filt_html(Left(OB_IIF(rs(2),"无题"),l)) & "</a>"
			If u=1 Then
			sAddon=OB_IIF(rs(4),"-")
			sAddon = "<span name=""nickname_"&rs("userid")&""" id=""nickname_"&rs("userid")&""">"&rs("userid")&"</span>"
			End if
			if t=1 Then
				If sAddon<>"" Then sAddon=sAddon & ","
				sAddon=sAddon & rs(3)
			End If
			If sAddon<>"" Then sAddon="(" & sAddon & ")"
			sRet=sRet & sAddon & "</li>" & Vbcrlf
			i = i + 1
			rs.Movenext
		Loop
		Set rs = Nothing
		sRet=sRet & "</ul>" & Vbcrlf
		sRet = sRet & oblog.GetNickNameById (arrayList,i,teamid&postnum&l&u&t)
	End if
	GetPosts=sRet
End Function

'最受欢迎的用户,计算方法
'user_siterefu_num+comment_count*1.5+message_count*1.5+sub_num*3
'访问数+回复数*1.5+留言数*1.5+被订阅数*3
't 是否显示用户头像
Function GetHotUsers(n,t)
	Dim rs, userurl,userico,i
	set rs=oblog.execute("select top "&n&" username,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder,user_icon1 from [oblog_user] where lockuser=0 and isdel=0 order by (user_siterefu_num+comment_count*1.5+message_count*1.5+sub_num*3) desc,userid DESC")
	GetHotUsers = Vbcrlf & "<ul>" & Vbcrlf
	While Not rs.EOF
		If oblog.cacheConfig(5) = 1 Then
			userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
		Else
			userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
		End If
		If t=1 Then userico="<img src=""" & OB_IIF(rs(8),"images/ico_default.gif") & """ width=""48"" height=""48"" border=""0"" /><br />"
		GetHotUsers=GetHotUsers&"<li><a href="&userurl&" target=""_blank"" title=""" & rs(2)& """>"&userico& rs(2)&"</a></li>" & vbcrlf
		rs.MoveNext
	Wend
	GetHotUsers=GetHotUsers&"</ul>" & Vbcrlf
	Set rs = Nothing
End Function

'随机调用博客链接,这里只调头像,博客名会默认显示在提示那里.
'show_rnduser(调用条数,图片高度,图片宽度,是否只调用推荐/活跃,多少天内登录过的活跃用户,是否只调用有自定义头像的用户)
'是否只调用推荐/活跃    1 只是推荐    10 只是推荐男生博客  11 只是推荐女生博客 2 按最后登录时间过滤 20按登录时间过滤男生  21按登录时间过滤女生
'是否只调用有自定义头像的用户 0 否 1 是
'$show_rnduser(40,48,48,2,30,1)$

Function GetRndUser(num,width,height,types,dht,ishaveface)
	Dim rs,sql,Utype,UFdate,RndOrderBy,userurl
	UFdate = int(dht)
	If Err Then Err.clear:UFdate = 30
		If Is_Sqldata = 1 Then
			RndOrderBy = " Order By Newid()"
		Else
			Randomize
			RndOrderBy = " Order By Rnd(-(UserID+"&Rnd()&"))"
		End If
		If ishaveface = "1" Then RndOrderBy=" and not(user_icon1 is null or user_icon1='') " & RndOrderBy
	Select Case types
		Case "1"
			Utype= " and user_isbest=1"
		Case "10"
			Utype= " and user_isbest=1 and sex=1"
		Case "11"
			Utype= " and user_isbest=1 and sex=0"
		Case "2"
			Utype= " and datediff("&G_Sql_d&",lastlogin,"&G_Sql_Now&") <= "&UFdate
		Case "20"
			Utype= " and sex=1 and datediff("&G_Sql_d&",lastlogin,"&G_Sql_Now&") <= "&UFdate
		Case "21"
			Utype= " and sex=0 and datediff("&G_Sql_d&",lastlogin,"&G_Sql_Now&") <= "&UFdate
	End Select
		Set rs=oblog.execute("select top "&num&" username,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder,user_icon1 from [oblog_user] where lockuser=0 and isdel=0 and (blog_password is null or blog_password='')  "&Utype&" "&RndOrderBy)
		GetRndUser = Vbcrlf & "<ul id=""showrnduser"">" & Vbcrlf
	While Not rs.EOF
		If oblog.cacheConfig(5) = 1 Then
			userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
		Else
			userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
		End If
			GetRndUser=GetRndUser&"<li><a href="""&userurl&""" target=""_blank"" title=""" & rs(2)& """><img src=""" & OB_IIF(rs(8),"images/ico_default.gif") & """ width="""&width&""" height=""" & height & """ border=""0"" title="""&rs(2)&"""/></a></li>" & vbcrlf
		rs.MoveNext
	Wend
	GetRndUser=GetRndUser&"</ul>" & Vbcrlf
	Set rs = Nothing


End Function

function TreeClass(n)
	dim Table_Name,wsql,toptitle,fname
	select case n
		case "user"
			Table_Name="oblog_userclass"
			wsql=""
			toptitle="用户类别"
			fname="listblogger.asp?usertype="
		case "log"
			Table_Name="oblog_logclass"
			wsql=" where idType=0 "
			toptitle="日志类别"
			fname="list.asp?classid="
		case "photo"
			Table_Name="oblog_logclass"
			wsql=" where idType=1 "
			toptitle="相片类别"
			fname="photo.asp?classid="
		case "group"
			Table_Name="oblog_logclass"
			toptitle=oblog.CacheConfig(69)& "类别"
			wsql=" where idType=2 "
			fname="groups.asp?classid="
	end select
	dim sqlClass,rsClass,D_String
	sqlClass="select id,parentid,classname From "&Table_Name&wsql&"  order by RootID,OrderID"
	set rsClass=oblog.execute(sqlClass)
	'把查询到的内容存放到字符串里，在JS中调用该字符串
	do	while not rsClass.eof
		D_String=D_String&"|"&rsClass("id")&","&rsClass("parentid")&",<a href='"&fname & rsClass("id") & "'>"&rsClass("classname")&"</a>,0"
		rsClass.movenext
	loop
	TreeClass="<script src='inc/tree.js'></script><script language='javascript' type='text/javascript'>var J_String,J_First,J_Second;var i,j;d = new dTree('d');d.add(0,-1,'<strong>"&toptitle&"</strong>');J_String="""&D_String&""";J_First=J_String.split('|');for(i=0;i<J_First.length;i++){J_Second=J_First[i].split(',');d.add(J_Second[0],J_Second[1],J_Second[2],'',J_Second[3]);}document.write(d);</script>"
	set rsClass=nothing
end function

'获得辩论列表
'n:显示条数;
'l:字符数目;
's:显示类型,1最新/2参与人数最多
Function GetArgueList(n,l,s)
	Dim sRet,Sql,rs,sState
	If s="1" Then
		'最新
		Sql="select top " & n & " argueid,topic,a_ico,actions,actions1,actions2,actions3 From oblog_argue Where istate=2 Order By argueid Desc"
	Else
		'最热的
		Sql="select top " & n & " argueid,topic,a_ico,actions,actions1,actions2,actions3  From oblog_argue Where istate=2 Order By actions Desc"
	End If
	'Response.Write Sql
	Set rs=oblog.Execute(Sql)
	Do While Not rs.Eof
		sRet=sRet & "<li><a href=""bl.asp?cmd=show&blid=" & rs("argueid") & """ target=""_blank"">" & Left(rs("topic"),l) & "</a><br/><font color=""red"">正</font>&nbsp;" & rs("actions1") & "&nbsp;&nbsp;<font color=""blue"">反</font>&nbsp;" & rs("actions2") & "&nbsp;&nbsp;<font color=""green"">参与</font>&nbsp;" &  rs("actions") & "</li>"
		rs.Movenext
	Loop
	Set rs=Nothing
	GetArgueList=sRet
	sRet=""
End Function

Function GetTemplate(n)
	Dim sRet,Sql,rs
	sql="SELECT TOP "&n&" * FROM oblog_userskin WHERE ispass=1 ORDER BY Id DESC"
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open sql ,CONN,1,1
	If Not RS.Eof Then
		While Not rs.EOF
			sRet = sRet &"<!-- 最新模板 -->"&vbcrlf
			sRet = sRet &"<div id=""NewSkin"">"&vbcrlf
			sRet = sRet &"	<div class=""SkinImg""><a href=""showskin.asp?id="&rs("id")&""" target =""_blank""><img src="""&rs("skinpic")&""" alt="""&rs("userskinname")&""" /></a></div>"&vbcrlf
			sRet = sRet &"	<div class=""Skinname""><a href=""showskin.asp?id="&rs("id")&""" target =""_blank"">"&rs("userskinname")&"</a></div>"&vbcrlf
			sRet = sRet &"</div>"&vbcrlf
			sRet = sRet &"<!-- 最新模板 END -->"&vbcrlf
			rs.MoveNext
		Wend
	End If
	GetTemplate = sRet
	sRet = ""
End Function

Function GetAlbum(n,l)
	Dim sRet,Sql,rs
	Dim Imgsrc,Preimgsrc,fso
	Set fso = Server.CreateObject(oblog.CacheCompont(1))
	Sql = "SELECT TOP "&N&" c.photo_path,c.subjectid,c.subjectlognum,userid,subjectname FROM "
	Sql = Sql &" oblog_subject AS c "
	Sql = Sql &" WHERE c.subjecttype = 1 AND (c.ishide = 0 OR c.ishide IS NULL) and c.subjectlognum>0"
	If L = 0 Then
		Sql = Sql &" ORDER BY c.subjectid DESC"
	Else
		Sql = Sql &" ORDER BY c.views DESC,c.subjectid DESC"
	End If
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open sql ,CONN,1,1
	If Not RS.Eof Then
		sRet = "<!-- 相册标签 -->"&vbcrlf
		sRet = sRet &"<div id=""NewPhotoAlbum"">"&vbcrlf
		While Not rs.EOF
			Imgsrc=RS(0)
			If Not IsNull(Imgsrc) Then
				Preimgsrc=Replace(Imgsrc,Right(Imgsrc,3),"Jpg")
				Preimgsrc=Replace(Preimgsrc,Right(Preimgsrc,Len(Preimgsrc)-instrrev(Preimgsrc,"/")),"Pre"&Right(Preimgsrc,Len(Preimgsrc)-instrrev(Preimgsrc,"/")))
				If Not Preimgsrc="" And Not IsNull(Preimgsrc) Then
				If Not Fso.Fileexists(Server.Mappath(Preimgsrc)) Then
					Preimgsrc=Imgsrc
				End If
				End If
			End if
			sRet = sRet &"	<div class=""NewPhotoAlbum"">"&vbcrlf
			sRet = sRet &"		<div class=""NewPhotoAlbumImg""><a href=""go.asp?albumid="&rs(3)&""" target = ""_blank""><img src="""&Proico(Preimgsrc,4)&""" /></a></div>"&vbcrlf
			sRet = sRet &"		<div class=""NewPhotoAlbumName""><a href=""go.asp?albumid="&rs(3)&""" target = ""_blank"">"&rs("subjectname")&"</a></div>"&vbcrlf
			sRet = sRet &"	</div>"&vbcrlf
			RS.MoveNext
		Wend
		sRet = sRet &"</div>"&vbcrlf
		sRet = sRet &"<!-- 相册标签 END -->"&vbcrlf
	End If
	GetAlbum = sRet
	sRet = ""
End Function

Function GetPic(n,l)
	Dim sRet,Sql,rs
	Dim Imgsrc,Preimgsrc,fso
	Set fso = Server.CreateObject(oblog.CacheCompont(1))
	Sql = "SELECT TOP "&N&" photo_path,photo_title,fileid FROM oblog_album "
	Sql = Sql &" WHERE (ishide = 0 OR ishide IS NULL)"
	If L = 0 Then
		Sql = Sql &" ORDER BY photoID DESC"
	ElseIf l = 1 Then
		Sql = Sql &" ORDER BY views DESC,photoID DESC"
	Else
		Sql = Sql &" ORDER BY commentnum DESC,photoID DESC"
	End If
'	OB_DEBUG SQL,1
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open sql ,CONN,1,1
	If Not RS.Eof Then
		sRet = "<!-- 相片标签 -->"&vbcrlf
		sRet = sRet &"<div id=""NewPhoto"">"&vbcrlf
		While Not rs.EOF
			Imgsrc=RS(0)

			Preimgsrc=Replace(Imgsrc,Right(Imgsrc,3),"Jpg")
			Preimgsrc=Replace(Preimgsrc,Right(Preimgsrc,Len(Preimgsrc)-instrrev(Preimgsrc,"/")),"Pre"&Right(Preimgsrc,Len(Preimgsrc)-instrrev(Preimgsrc,"/")))
			If Not Fso.Fileexists(Server.Mappath(Preimgsrc)) Then
				Preimgsrc=Imgsrc
			End If
			sRet = sRet &"	<div class=""NewPhoto"">"&vbcrlf
			sRet = sRet &"		<div class=""NewPhotoImg""><a href=""go.asp?fileid="&rs(2)&""" target = ""_blank""><img src="""&Proico(Preimgsrc,4)&""" /></a></div>"&vbcrlf
			sRet = sRet &"		<div class=""NewPhotoName""><a href=""go.asp?fileid="&rs(2)&""" target = ""_blank"">"&OB_IIF(rs(1),"无标题")&"</a></div>"&vbcrlf
			sRet = sRet &"	</div>"&vbcrlf
			RS.MoveNext
		Wend
		sRet = sRet &"</div>"&vbcrlf
		sRet = sRet &"<!-- 相片标签 END -->"&vbcrlf
	End If
	GetPic = sRet
	sRet = ""
End Function

Function GetDiggs(n,l)
	Dim sRet,Sql,rs,ClassName
	Dim arrayList,i
	ReDim arrayList(n-1)
	Sql = "SELECT TOP "&N&" a.diggnum,a.diggurl,a.diggtitle,a.addtime,a.author,a.authorid FROM oblog_userdigg AS a INNER JOIN oblog_log AS c ON a.logid = c.logid WHERE a.istate = 1 AND c.isdel=0 "
	If L = 0 Then
		Sql = Sql &" ORDER BY a.DiggID DESC"
		ClassName = "NewDIGG"
	ElseIf l = 1 Then
		Sql = Sql &" ORDER BY a.diggnum DESC,a.DiggID DESC"
		ClassName = "DIGGTop"
	Else
		Sql = Sql &" ORDER BY a.lastdiggtime DESC"
	End If
'	OB_DEBUG SQL,1
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open sql ,CONN,1,1
	If Not RS.Eof Then
		i = 0
		sRet = "<!-- DIGG标签 -->"&vbcrlf
		sRet = sRet &"<div id="""&ClassName&""">"&vbcrlf
		While Not rs.EOF
			arrayList(i) = rs("authorid")
			sRet = sRet &"	<div class="""&ClassName&""">"&vbcrlf
			sRet = sRet &"		<span class=""DIGGNumber"">"&rs(0)&"</span>"&vbcrlf
			sRet = sRet &"		<span class=""DIGGTitle""><a href="""&rs(1)&""" title="""&rs(2)&""">"&rs(2)&"</a></span>"&vbcrlf
			If l = 0 Then
				sRet = sRet &"		<span class=""DIGGTime"">"&rs(3)&"</span>"&vbcrlf
				sRet = sRet &"		<span class=""DIGGUser""><a href=""go.asp?userid="&rs(5)&"""><span name=""nickname_"&rs("authorid")&""" id=""nickname_"&rs("authorid")&""">"&rs("authorid")&"</span></a></span>"&vbcrlf
			End If
			sRet = sRet &"	</div>"&vbcrlf
			i = i + 1
			RS.MoveNext
		Wend
		sRet = sRet &"</div>"&vbcrlf
		sRet = sRet &"<!-- DIGG标签 END -->"&vbcrlf
		sRet = sRet & oblog.GetNickNameById (arrayList,i,n&l)
	End If
	GetDiggs = sRet
	sRet = ""
End Function

Function GetUserDiggs(n,l)
	Dim sRet,Sql,rs
	Sql = "SELECT TOP "&N&" userid,User_Icon1,username,nickname,diggs FROM "
	Sql = Sql &" oblog_user "
	Sql = Sql &" WHERE lockuser=0 AND isdel=0  "
	If L = 0 Then
		Sql = Sql &" ORDER BY diggs DESC,userid DESC"
	Else
		Sql = Sql &" ORDER BY userid DESC"
	End If
'	OB_DEBUG SQL,1
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open sql ,CONN,1,1
	If Not RS.Eof Then
		sRet = "<!-- DIGG标签 ,控制图片大小请使用CSS控制 DIGGMostUserIco 里的 img 属性-->" & vbcrlf
		sRet = sRet &"<div id=""DIGGMostUser"">" & vbcrlf
		While Not rs.EOF
			sRet = sRet &"	<div class=""DIGGMostUser"">" & vbcrlf
			sRet = sRet &"		<div class=""DIGGMostUserIco""><a href=""go.asp?userid="&rs(0)&""" target = ""_blank""><img src="""&Proico(rs(1),1)&""" alt="""&OB_IIF(rs(3),rs(2))&""" /></a></div>" & vbcrlf
			sRet = sRet &"		<div class=""DIGGMostUserName""><a href=""go.asp?userid="&rs(0)&""" title=""alt="""&OB_IIF(rs(3),rs(2))&""""" target = ""_blank"">"&OB_IIF(rs(3),rs(2))&"</a>被推荐<span title="""&OB_IIF(rs(4),0)&""">"&OB_IIF(rs(4),0)&"</span>次</div>" & vbcrlf
			sRet = sRet &"	</div>" & vbcrlf
			RS.MoveNext
		Wend
		sRet = sRet &"</div>" & vbcrlf
		sRet = sRet &"<!-- DIGG标签 END -->" & vbcrlf
	End If
	GetUserDiggs = sRet
	sRet = ""
End Function

%>
