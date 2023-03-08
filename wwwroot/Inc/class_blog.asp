<!--#include file="../inc/Inc_UBB.asp"-->
<!--#include file="../inc/Inc_Tags.asp"-->
<!--#include file="../inc/Cls_XmlDoc.asp"-->
<%
Class Class_Blog
	Public GoUrl, user_skin_main, user_skin_showlog, user_userName, user_id, user_nickName, user_showName
	Public user_commentasc, user_path, user_folder,user_showlog_num, user_showlogword_num, BlogName, user_siteinfo,user_truepath,user_trueurl,user_group,PassPort_userid,PlayerType
	Public user_Blogpassword, user_domain, user_placard, user_links, user_log_count, user_comment_count,user_indexlist
	Public user_message_count, user_shownewlog_num, user_shownewmessage_num, user_shownewcomment_num,log_truepath,user_level,user_subnum
	Public rs, objFSO, tf, ispwBlog,showpwblog,showpwlog,Page,user_province,user_city
	Public m_index,m_log,m_subjectid,m_subjectindex,m_message,m_album,m_info,m_placard,m_links,m_newblog,m_newmessage,m_comment,m_subject,m_subject_l,m_commentsmore,m_friends,m_mygroups
	Public Cache_Name
	Private Sub Class_Initialize()
		Set objFSO = Server.CreateObject(oblog.CacheCompont(1))
		showpwblog=False
		showpwlog=False
		Cache_Name=cookies_name
	End Sub

	Private Sub Class_Terminate()
		Set objFSO = Nothing
		Set tf = Nothing
		Set rs = Nothing
	End Sub

	Public Property Let userid(Byval Values)
		dim rstmp,strSql
		userid = Int(Values)
		strSql = "select user_dir,user_showlog_num,user_showlogword_num,user_skin_main,user_skin_showlog,"
		strSql = strSql & "BlogName,nickName,userName,siteinfo,Blog_password,"
		strSql = strSql & "comment_isasc,user_domain,user_domainroot,user_placard,user_links,"
		strSql = strSql & "log_count,comment_count,message_count,user_shownewlog_num,user_shownewcomment_num,"
		strSql = strSql & "user_shownewmessage_num,user_folder,user_level,province,city,sub_num,user_group,"
		strSql = strSql & "passport_userid,PlayerType,indexlist"&str_domain
		strSql = strSql & " From oBlog_user Where userid=" & userid
		Set rs = oblog.Execute(strSql)
		If rs.EOF Then Exit Property
		user_id = userid
		user_path = Trim(rs("user_dir")) & "/" & rs("user_folder")
		user_showlog_num = Int(rs("user_showlog_num"))
		G_P_PerMax=user_showlog_num
		user_showlogword_num = rs("user_showlogword_num")
		user_skin_main = Trim(rs("user_skin_main"))
		user_skin_showlog = Trim(rs("user_skin_showlog"))
		BlogName = oblog.filt_html(rs("BlogName"))
		user_nickName = oblog.filt_html(rs("nickName"))
		user_userName = oblog.filt_html(rs(7))
		user_siteinfo = oblog.filt_html(rs(8))
		user_Blogpassword = Trim(rs(9))
		user_commentasc = rs(10)
		user_domain = Trim(rs(11)) & "." & Trim(rs(12))
		user_placard=rs(13)
		user_links=rs(14)
		user_log_count=rs(15)
		user_comment_count=rs(16)
		user_message_count=rs(17)
		user_shownewlog_num=rs(18)
		user_shownewcomment_num=rs(19)
		user_shownewmessage_num=rs(20)
		user_folder=rs("user_folder")
		user_level=rs("user_level")
		user_province=rs("province")
		user_city=rs("city")
		user_subnum=rs("sub_num")'订阅数
		user_indexlist=rs("indexlist")
		user_group = rs("user_group")
		PassPort_userid = rs("passport_userid")
		PlayerType = rs("PlayerType")
		'判断是否真实域名
		if true_domain=1 then
			if rs("custom_domain")<>"" and not isnull(rs("custom_domain")) then
				user_domain=rs("custom_domain")
			end if
			user_truepath="http://"&user_domain&"/"
			user_trueurl=user_truepath & "index." & f_ext
			log_truepath=""
		else
			user_truepath=blogdir&user_path&"/"
			user_trueurl=oblog.cacheConfig(3) &  user_path & "/index." & f_ext
			If oblog.CacheConfig(4) <>"" And oblog.CacheConfig(5) = 1 Then
				user_trueurl = "http://"&user_domain&"/"
			End if
			log_truepath=blogdir
		end If
		If user_skin_main = "" Or IsNull(user_skin_main) Or IsNull(user_skin_showlog) Or user_skin_showlog=""  Then

			user_skin_main = Application(Cache_Name & "_user_skin_main")
			user_skin_showlog = Application(Cache_Name & "_user_skin_showlog")
			If user_skin_main = "" Or IsNull(user_skin_main) Or IsNull(user_skin_showlog) Or user_skin_showlog=""  Then
				Set rstmp = oblog.Execute("select skinmain,skinshowlog from oBlog_userskin where isdefault=1")
				If Not rstmp.EOF Then
					Application.Lock
					Application(Cache_Name & "_user_skin_main") = rstmp(0)
					Application(Cache_Name & "_user_skin_showlog")  = rstmp(1)
					Application.unLock
					user_skin_main = rstmp(0)
					user_skin_showlog = rstmp(1)
				Else
					Set rstmp = Nothing
					Set rs = Nothing
					Response.Write ("模版错误")
					Response.End
				End If
			End If

		End If
	  '----------------------------------------------------------------
		If user_Blogpassword = "" Or IsNull(user_Blogpassword) Then ispwBlog = False Else ispwBlog = True
	End Property

	public sub update_log(logid,resp)
		Dim vote0,vote1,sTeamAddon, rst
		Dim sql, rstmp,user_path_new,user_logpath_new,user_domain_new,user_userName_new,user_nickname_new,user_skin_main_new,user_skin_showlog_new,bTeam
		Dim show_topic, show_emot, show_addtime, show_logtext, show_author, show_loginfo, show_more, show_logcyc, show_topictxt, show_logmore, show, log_month, user_logpath1, log_title, commentasc,faceurl
		Dim homepage_str, commentid, commenttopic, strtmp, encommment, i, filename, injsfile,user_logpath,logtype
		bTeam=false
		If bTeam=false Then
			user_path_new=user_path
			user_domain_new=user_domain
			user_userName_new=user_username
			user_nickname_new=user_nickname
			user_skin_main_new=user_skin_main
			user_skin_showlog_new=user_skin_showlog
		End If
		logid = Int(logid)
		Set rs = oblog.Execute("select face,topic,logtext,author,istop,isencomment,addtime,ishide,ispassword,isbest,commentnum,trackbacknum,passcheck,authorid,filename,logtype,vote1,vote0,userid,isneedlogin,viewscores,Abstract,isspecial,viewgroupid from oblog_log where isdel=0 and logid=" & logid)
		If rs.EOF Then Exit Sub
		If rs("userid")<>rs("authorid") Then
			bTeam=true
			Set rst=oblog.Execute("select userid,username,nickName,user_domain,user_domainroot,user_folder,user_dir,user_skin_main,user_skin_showlog From oblog_user Where userid=" & rs("userid"))
			If rst.Eof Then
				Exit Sub
				Set rst=Nothing
			End If
			user_path_new = Trim(rst("user_dir")) & "/" & rst("user_folder")
			user_domain_new =  Trim(rst("user_domain")) & "." & rst("user_domainroot")
			user_userName_new=rst("username")
			user_nickname_new=OB_IIF(rst("nickname"),user_userName_new)
			'user_skin_main_new=FilterJs(rst("user_skin_main"))
			'user_skin_showlog_new=FilterJs(rst("user_skin_showlog"))
			user_skin_main_new=OB_IIF(rst("user_skin_main"),"请重新选择模板")
			user_skin_showlog_new=OB_IIF(rst("user_skin_showlog"),"请重新选择模板")
			Set rst=Nothing
		End If
		If Int(Month(rs("addtime"))) < 10 Then
			log_month = Year(rs("addtime")) & "0" & Month(rs("addtime"))
		Else
			log_month = Year(rs("addtime")) & Month(rs("addtime"))
		End If
		if oblog.CacheConfig(57)="0" then
			user_logpath=user_path_new
		else
			user_logpath=user_path_new&"/archives/"&Trim(year(rs("addtime")))
		end if
		logtype=rs("logtype")
		filename = Trim(rs("filename"))
		If filename = "" Or IsNull(filename) Then filename = logid
		encommment = rs("isencomment")
		strtmp = ""
		If rs("passcheck") = 0 Then
			strtmp = "此日志需要管理审核后才可浏览。"
		Else
			If Not showpwlog Then
				If rs("ishide") = 1 Then
					strtmp = "此日志为隐藏日志，仅好友可浏览，<a href=""" & blogurl & "more.asp?id=" & logid & """>点击进入验证页面</a>。"
				ElseIf  rs("ispassword") <> "" Then
					strtmp = "<form method='post' action='" & blogurl & "more.asp?id=" & logid & "'>请输入日志访问密码：<input type=""password"" size=""15"" name=""password"" />  <input type=""submit"" value=""提交""></form>"
				ElseIf OB_IIF(rs("isneedlogin"),0) = 1  Then
					strtmp =oblog.filt_html_b(rs("Abstract"))	&"<br /><br />此日志需要登录后才可浏览，<a href='" & blogurl & "more.asp?id=" & logid & "'>点击进入验证页面</a>。"
				ElseIf OB_IIF(rs("viewscores"),0) > 0 Then
					strtmp = oblog.filt_html_b(rs("Abstract"))	&"<br /><br />此日志需要消费积分<strong>"&rs("viewscores")&"</strong>才可浏览，<a href='javascript:void(null);' onclick=""if(confirm('浏览此日志需消费积分"&rs("viewscores")&"，确认浏览？')==true)self.location='" & blogurl & "more.asp?id=" & logid & "';return false;"">点击进入验证页面</a>。"
				ElseIf OB_IIF(rs("viewgroupid"),0) > 0 Then
					strtmp =oblog.filt_html_b(rs("Abstract"))	&"<br /><br />此日志需要特定用户组才可浏览，<a href='" & blogurl & "more.asp?id=" & logid & "'>点击进入验证页面</a>。"
				Else
					If OB_IIF(rs("isspecial"),0) > 0 Then
						strtmp = "此日志为特殊日志，<a href='" & blogurl & "more.asp?id=" & logid & "'>点击进入验证页面</a>"
					End if
				End If
			End if
		End if
		If user_nickName_new <> "" Then user_showName = user_nickName_new Else user_showName = user_userName_new
		'If rs("face") = "0" Then show_emot = "" Else show_emot = "<img src=""" & blogurl & "images/face/" & rs("face") & ".gif"" />"
		show_topictxt = oblog.filt_html(OB_IIF(rs("topic"),"无标题"))
		log_title = show_topictxt
		commenttopic = "Re:" & show_topictxt
		If rs("isbest") = 1 Then show_topictxt = show_topictxt & "　<img src=""" & blogurl & "images/jhinfo.gif"" />"
		'show_topic = show_emot
		show_addtime = rs("addtime")
		show_topic = show_topic & show_topictxt
		If user_nickName = "" Or IsNull(user_nickName) Then
			show_author = user_userName_new
		Else
			show_author = user_nickName_new
		End If
		If rs("authorid") <> user_id Then show_author = rs("author")
		show_loginfo = show_author & " 发表于 " & show_addtime
		show_more = "<a href=""#"" >阅读全文<span id=""ob_logreaded""></span></a>"
		show_more = show_more & " | " & "<a href=""#cmt"">回复(" & rs("commentnum") & ")</a> <span id = ""ob_logm"&logid&"""> </span>"
		show_more = show_more & " | <a href=""" & blogurl & "showtb.asp?id=" & logid & """ target=""_blank"">引用通告<span id=""ob_tbnum""></span></a>"
		injsfile = "<Script src=""" & blogurl & "count.asp?action=logtb31&id=" & logid & """></Script>"
		show_more = show_more & " | <a href=""" & blogurl & "user_post.asp?logid=" & logid & """ target=""_blank"">编辑</a>"
		If strtmp <> "" Then
			show_logtext = strtmp
		Else
			show_logtext = ob_IIF(rs("logtext"),"未输入内容.")
			If Left(show_logtext, 7) = "#isubb#" Then
				show_logtext = UBBCode(show_logtext, 1)
				show_logtext = Replace(show_logtext, Chr(10), "<br /> ")
				'show_logtext=oblog.filt_html_b(show_logtext)
			End If
			show_logtext = Replace(show_logtext, "#isubb#", "")
			show_logtext = filtimg(show_logtext)
			Dim showDes
			showDes = show_logtext
		End If
		show_logtext = "<span id=""ob_logd"&logid&"""></span> " & show_logtext
		'-----------------------------Addon Start--------------------
		Dim sAddon,sAddOn1,sAddon2
		'标签
		sAddon1=Tags_ShowForBlog(logid,user_truepath)
		'群组信息
		Set rst=oblog.Execute("select a.teamid,a.t_name From oblog_team a,oblog_teampost b Where a.teamid=b.teamid And b.logid=" & logid)
		Do While Not rst.Eof
			sAddon2=sAddon2 & "<span><a href="""&blogurl&"group.asp?gid=" & rst(0) & """ target=_blank>" & rst(1) & "</a></span>&nbsp;"
			rst.Movenext
		Loop
		'OB_Debug sAddon2,1
		Set rst=Nothing
		If sAddon1&sAddon2<>""  Then
			sAddon="<div id=""blogaddon"">" & vbcrlf
			If sAddon1<>"" Then sAddon=sAddon & sAddon1
			If sAddon2<>"" Then	sAddon=sAddon & "<li>" &oblog.CacheConfig(69)& "：" & sAddon2&"</li>" & vbcrlf
			sAddon=sAddon & "</div>" & vbcrlf
		End if
		'-----------------------------Addon End--------------------
		show_logcyc = Replace(user_skin_showlog, "$show_topic$", show_topic)
		show_logcyc = Replace(show_logcyc, "$show_loginfo$", show_loginfo)
		show_logcyc = Replace(show_logcyc, "$show_logtext$", show_logtext)
		show_logcyc = Replace(show_logcyc, "$show_more$", show_more)
		show_logcyc = Replace(show_logcyc, "$show_emot$", show_emot)
		show_logcyc = Replace(show_logcyc, "$show_author$", show_author)
		show_logcyc = Replace(show_logcyc, "$show_addtime$", show_addtime)
		show_logcyc = Replace(show_logcyc, "$show_topictxt$", show_topictxt)

		show_logcyc = Replace(show_logcyc, "$show_blogtag$", "")

		show_logcyc = Replace(show_logcyc, "$show_blogzhai$", "<div id=""blogzhai""></div>")
		show_logmore = show_logcyc
		show_logmore=show_logmore&sAddon
		show_logmore = show_logmore & vbcrlf & "<div id=""morelog"">" & vbcrlf
		show_logmore = show_logmore&"	<ul>" & vbcrlf
		set rstmp=oblog.execute("select top 1 logfile,topic from oblog_log where logid<"&logid&" and userid="&user_id&" and logtype="&logtype&" and isdraft=0 and isdel=0 order by addtime desc")
		if not rstmp.eof then
			show_logmore = show_logmore&"		<li>上一篇："&"<a href="""&log_truepath&rstmp(0)&""">"&oblog.filt_html(rstmp(1))&"</a></li>" & vbcrlf
			rstmp.movenext
		end if
		set rstmp=oblog.execute("select top 1 logfile,topic from oblog_log where logid>"&logid&" and userid="&user_id&" and logtype="&logtype&" and isdraft=0 and isdel=0 order by addtime asc")
		if not rstmp.eof then
			show_logmore = show_logmore&"		<li>下一篇："&"<a href="""&log_truepath&rstmp(0)&""">"&oblog.filt_html(rstmp(1))&"</a></li>" & vbcrlf
			rstmp.movenext
		end if
		show_logmore = show_logmore&"	</ul>" & vbcrlf
		show_logmore = show_logmore&"</div>" & vbcrlf
		'vote0=OB_IIF(rs("vote0"),"0")
		'vote1=OB_IIF(rs("vote1"),"0")
		If strtmp = "" Then
			If user_commentasc = 1 Then commentasc = " order by commentid asc" Else commentasc = " order by commentid desc"
			Set rs = oblog.Execute("select top 40 comment_user,commenttopic,comment,addtime,commentid,homepage,isguest,ubbedit from oblog_comment where istate =1 and isdel=0 and mainid=" & logid & commentasc)
			If Not rs.EOF Then
				While Not rs.EOF
					If IsNull(rs(5)) Then
						homepage_str = "个人主页"
					Else
						If Trim(Replace(rs(5), "http://", "")) = "" Then
							homepage_str = "个人主页"
						Else
							homepage_str = "<a href=""" &blogurl&"go.asp?url="& oblog.filt_html(rs(5)) & """ target=""_blank"">个人主页</a>"
						End If
					End If
					commentid = rs(4)
					show_topic = oblog.filt_html(rs(1)) & "<a name='" & rs(4) & "'></a>"
					If rs(6) = 1 Then
						show_author = "<span id=""n_" & commentid & """>" & oblog.filt_html(rs(0)) & "(游客)</span>"
						faceurl=blogurl&"images/ico_default.gif"
					Else
						show_author = "<span id=""n_" & commentid & """>" & oblog.filt_html(rs(0)) & "</span>"
						set rstmp=oblog.execute("select user_icon1 from oblog_user where username='"&oblog.filt_badstr(rs(0))&"'")
						if not rstmp.eof then
							faceurl = ProIco (rstmp(0),1)
						else
							faceurl = blogurl&"images/ico_default.gif"
						end if
					End If
					faceurl="<img class=""ob_face"" src="""&faceurl&""" width=""48"" height=""48"" align=""absmiddle"" alt=""" & oblog.filt_html(rs(0))
					If rs(6) = 1 Then
						faceurl = faceurl & "(游客)"
					End If
					faceurl= faceurl & """ />"

					faceurl=Replace(homepage_str,"个人主页",faceurl)
					show_addtime = "<span id=""t_" & commentid & """>" & rs(3) & "</span>"
					show_topictxt = OB_IIF(show_topic,"无题")
					show_loginfo = show_author & "发表评论于" & show_addtime
					show_logtext = faceurl&"<span id=""c_" & commentid & """>"
					If rs("ubbedit")= 2 Then
						show_logtext = show_logtext & oblog.FilterUbbFlash(filtscript(rs(2)))
					Else
						show_logtext = show_logtext & oblog.Ubb_Comment(rs(2))
					End if
					show_logtext = show_logtext &"</span>"
					show_more = homepage_str & " | <a href=""javascript:reply_quote('" & commentid & "')"" >引用</a> | <a href=""#top"">返回</a>"
					show_more = show_more & " | <a href=""" & blogurl & "user_comments.asp?action=del&id=" & commentid & """  target=""_blank"">删除</a>"
					show_more = show_more & " | <a href=""" & blogurl & "user_comments.asp?action=modify&re=true&id=" & commentid & """  target=""_blank"">回复</a>"
					show_logcyc = Replace(user_skin_showlog, "$show_topic$", show_topic)
					show_logcyc = Replace(show_logcyc, "$show_loginfo$", show_loginfo)
					show_logcyc = Replace(show_logcyc, "$show_logtext$", show_logtext)
					show_logcyc = Replace(show_logcyc, "$show_more$", show_more)
					show_logcyc = Replace(show_logcyc, "$show_emot$", "")
					show_logcyc = Replace(show_logcyc, "$show_author$", show_author)
					show_logcyc = Replace(show_logcyc, "$show_addtime$", show_addtime)
					show_logcyc = Replace(show_logcyc, "$show_topictxt$", show_topictxt)
					show_logmore = show_logmore & show_logcyc
					show_logmore = Replace(show_logmore, "$show_blogtag$", "")
					show_logmore = Replace(show_logmore, "$show_blogzhai$", "")
					rs.movenext
					i = i + 1
				Wend
			End If

			If i >= 40 Then
				show_logmore = show_logmore & "<div id=""saveurl""><a href=""" & blogurl & "more.asp?action=comment&id=" & logid & "&page=1"">查看所有评论</a></div>"
			End If
			'Ajax Mode
			'show_logmore = show_logmore & "<div id=""saveurl""> ::<a href=""javascript:SendRequest('" & blogurl & "AjaxServer.asp?action=vote&v=1&logid=" & logid & "','ob_log_msg','');"">"&C_Vote_Action1&"("&vote1&")</a>::"
			'show_logmore = show_logmore & "<a href=""javascript:SendRequest('" & blogurl & "AjaxServer.asp?action=vote&v=0&logid=" & logid & "','ob_log_msg','');"">"&C_Vote_Action2&"("&vote0&")</a>::</div>"
			'show_logmore = show_logmore & "<div id=""ob_log_msg""></div>"
			If encommment = 1 Then
				Dim strguest
				If oblog.cacheConfig(27) = 1 Then strguest = "(游客无须输入密码)" Else strguest = ""
				show_logmore = filt_inc(show_logmore)
				show_logmore = show_logmore & vbCrLf & "<div id=""form_comment"">" & vbCrLf
				show_logmore = show_logmore & "	#gg_usercomment#<a name=""cmt""></a><div class=""title"">发表评论：</div>" & vbCrLf
				show_logmore = show_logmore & "	<form action=""" & blogurl & "savecomment.asp?logid=" & logid & """ method=""post"" name=""commentform"" id=""commentform"" onSubmit=""return Verifycomment()"">" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d1""><label>昵称：<input name=""UserName"" type=""text"" id=""UserName"" size=""20"" maxlength=""20"" value="""" /></label></div>" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d2""><label>密码：<input name=""Password"" type=""password"" id=""Password"" size=""20"" maxlength=""20"" value="""" /> " & strguest & "</label></div>" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d3""><label>主页：<input name=""homepage"" type=""text"" id=""homepage"" size=""42"" maxlength=""50"" value=""http://"" /></label></div>" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d4""><label>标题：<input name=""commenttopic"" type=""text"" id=""commenttopic"" size=""42"" maxlength=""50"" value=""" & commenttopic & """ /></label></div>" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d5"">" & vbCrLf
				show_logmore = show_logmore & "			<input type=""hidden"" name=""edit"" id=""edit"" value="""" />" & vbCrLf
				show_logmore = show_logmore & "			<div id=""oblog_edit"">"& oblog.CacheConfig(41)&"</div>" & vbCrLf
				show_logmore = show_logmore & "		</div>" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d6""><span id=""ob_code""></span><input type=""submit"" value=""&nbsp;提&nbsp;交&nbsp;"" onclick='oblog_edittext.createTextRange().execCommand(""Copy"");'></div>" & vbCrLf
				show_logmore = show_logmore & "	</form>" & vbCrLf
				show_logmore = show_logmore & "</div>" & vbCrLf
			End If
		End If
		show = Replace(user_skin_main_new, "$show_log$", show_logmore)
		If showpwblog = False And showpwlog = False Then
			show = repl_label(show, injsfile, log_title & "--" & BlogName, user_userName_new & "," & user_nickName_new, log_title, Left(RemoveHTML(showDes), 80), log_month)
			if true_domain=1 then
				user_logpath1 = Replace(user_logpath,user_path_new,"http://"&user_domain_new) & "/" & filename & "." & f_ext
			else
				user_logpath1 = user_logpath & "/" & filename & "." & f_ext
			end if
			show=Replace(show,"$show_calendar$","<!-- #include file=""..\..\calendar\"&log_month&".htm"" -->")
			If ispwblog = False Then
				savefile user_logpath,"\"&filename&"."&f_ext,show
			Else
				'savefile user_logpath,"\"&filename&"."&f_ext,"<script language=javascript>window.location.replace('"&blogurl&"pwblog.asp?action=log&userid="&user_id&"&logid="&logid&"')</script>"
				savefile user_logpath,"\"&filename&"."&f_ext,"<script language=javascript>window.location.replace('"&blogurl&"more.asp?id="&logid&"')</script>"
			End If
			oblog.execute("update oblog_log set logfile='"&user_logpath1&"' where logid="&logid)
			If resp = 1 Then
				gourl = user_logpath1
			ElseIf resp = 2 Then
				Response.Redirect (user_logpath1)
			ElseIf resp = 3 Then
				gourl = user_logpath1
			End If
		Else
			If f_ext = "htm" Or f_ext = "html" Then
				m_log=Replace(show,"$show_calendar$","<div id=""calendar""></div><script src='"&user_path_new&"/calendar/"&log_month&".htm'></script>")
			Else
				m_log=Replace(show,"$show_calendar$","<div id=""calendar"">"&oblog.readfile(user_path_new&"\calendar",log_month&".htm")&"</div>")
			End If
			m_log=m_log&injsfile
		End If
	End Sub

	public sub showcmt(logid)
		dim sql,rstmp
		dim show_topic,show_emot,show_addtime,show_logtext,show_author,show_loginfo,show_more,show_logcyc,show_topictxt,show_logmore,show,commentasc,faceurl
		dim homepage_str,commentid,strtmp
		logid=Int(logid)
		if user_commentasc=1 then commentasc=" order by commentid asc"	else commentasc=" order by commentid desc"
			set rs=Server.CreateObject("Adodb.RecordSet")
			rs.open "select comment_user,commenttopic,comment,addtime,commentid,homepage,isguest,ubbedit from oblog_comment where istate= 1 and isdel=0 and  mainid="&logid&commentasc,conn,1,1
			if rs.eof and rs.bof then
				show_logmore=show_logmore & "共有0篇评论<br>"
			else
				dim show_page,i
				G_P_FileName="more.asp?action=comment&id="&logid
				G_P_AllRecords=rs.recordcount
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
				if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
					rs.move (G_P_This-1)*G_P_PerMax
					show_page=oblog.showpage(false,true,"篇日志")
				end if
				do while not rs.eof
					if isnull(rs(5)) then
						homepage_str="个人主页"
					else
						if Trim(Replace(rs(5),"http://",""))="" then
							homepage_str="个人主页"
						else
							homepage_str="<a href=""" &blogurl&"go.asp?url="&oblog.filt_html(rs(5))&""" target=""_blank"">个人主页</a>"
						end if
					end If
					commentid=rs(4)
					show_topic=oblog.filt_html(rs(1))&"<a name='"&rs(4)&"'></a>"
					if rs(6)=1 then
						show_author="<span id=""n_"&commentid&""">"&oblog.filt_html(rs(0))&"(游客)</span>"
						faceurl=blogurl&"images/ico_default.gif"
					else
						show_author="<span id=""n_"&commentid&""">"&oblog.filt_html(rs(0))&"</span>"
						set rstmp=oblog.execute("select user_icon1 from oblog_user where username='"&oblog.filt_badstr(rs(0))&"'")
						if not rstmp.eof then
							faceurl = ProIco (rstmp(0),1)
						else
							faceurl=blogurl&"images/ico_default.gif"
						end if
					end If
					faceurl="<img class=""ob_face"" src="""&faceurl&""" width=""48"" height=""48"" align=""absmiddle"" />"
					faceurl=Replace(homepage_str,"个人主页",faceurl)
					show_addtime="<span id=""t_"&commentid&""">"&rs(3)&"</span>"
					show_topictxt=show_topic
					show_loginfo=show_author&"发表评论于"&show_addtime
					show_logtext = faceurl &"<span id=""c_" & commentid & """>"
					If rs("ubbedit")= 2 Then
						show_logtext = show_logtext & oblog.FilterUbbFlash(filtscript(rs(2)))
					Else
						show_logtext = show_logtext & oblog.Ubb_Comment(rs(2))
					End if
					show_logtext = show_logtext &"</span>"
					show_more=homepage_str&" | <a href=""javascript:reply_quote('"&commentid&"')"" >引用</a> | <a href=""#top"">返回</a>"
					show_more=show_more&" | <a href=""user_comments.asp?action=del&id="&commentid&"""  target=""_blank"">删除</a>"
					show_logcyc=Replace(user_skin_showlog,"$show_topic$",show_topic)
					show_logcyc=Replace(show_logcyc,"$show_loginfo$",show_loginfo)
					show_logcyc=Replace(show_logcyc,"$show_logtext$",show_logtext)
					show_logcyc=Replace(show_logcyc,"$show_more$",show_more)
					show_logcyc=Replace(show_logcyc,"$show_emot$","")
					show_logcyc=Replace(show_logcyc,"$show_author$",show_author)
					show_logcyc=Replace(show_logcyc,"$show_addtime$",show_addtime)
					show_logcyc=Replace(show_logcyc,"$show_topictxt$",show_topictxt)
					show_logmore=show_logmore&show_logcyc
					i=i+1
					if i>=G_P_PerMax then exit do
					rs.movenext
				loop
			end if
			show_logmore = Replace(show_logmore, "$show_blogtag$", "")
			show_logmore = Replace(show_logmore, "$show_blogzhai$", "")
			show_logmore=show_logmore&show_page
			if encommment=1 then
				dim strguest
				If oblog.cacheConfig(27) = 1 Then strguest = "(游客无须输入密码)" Else strguest = ""
				show_logmore = filt_inc(show_logmore)
				show_logmore = show_logmore & vbCrLf & "<div id=""form_comment"">" & vbCrLf
				show_logmore = show_logmore & "	#gg_usercomment#<a name=""comment""></a><div class=""title"">发表评论：</div>" & vbCrLf
				show_logmore = show_logmore & "	<form action=""" & blogurl & "savecomment.asp?logid=" & logid & """ method=""post"" name=""commentform"" id=""commentform"" onSubmit=""return Verifycomment()"">" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d1""><label>昵称：<input name=""UserName"" type=""text"" id=""UserName"" size=""20"" maxlength=""20"" value="""" /></label></div>" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d2""><label>密码：<input name=""Password"" type=""password"" id=""Password"" size=""20"" maxlength=""20"" value="""" /> " & strguest & "</label></div>" & vbCrLf
				show_logmore = show_logmore & "		<div class=""d3""><label>主页：<input name=""homepage"" type=""text"" id=""homepage"" size=""42"" maxlength=""50"" value=""http://"" /></label></div>"
				show_logmore = show_logmore & "		<div class=""d4""><label>标题：<input name=""commenttopic"" type=""text"" id=""commenttopic"" size=""42"" maxlength=""50"" value=""" & commenttopic & """ /></label></div>"
				show_logmore = show_logmore & "		<div class=""d5"">" & vbCrLf
				show_logmore = show_logmore & "			<input type=""hidden"" name=""edit"" id=""edit"" value="""" />" & vbCrLf
				show_logmore = show_logmore & "			<div id=""oblog_edit""></div>" & vbCrLf
				show_logmore = show_logmore & "		</div>" & vbCrLf
				'show_logmore = show_logmore & "		<div class=""d6""><script src=""" & blogurl & "count.asp?action=code""></script><input type=""submit"" value=""&nbsp;提&nbsp;交&nbsp;"" onclick='oblog_edittext.createTextRange().execCommand(""Copy"");'></div>" & vbCrLf

			show_logmore = show_logmore & "		<div class=""d6""><span id=""ob_code""></span><input type=""submit"" value=""&nbsp;提&nbsp;交&nbsp;"" onclick='oblog_edittext.createTextRange().execCommand(""Copy"");'></div>" & vbCrLf
				show_logmore = show_logmore & "	</form>" & vbCrLf
				show_logmore = show_logmore & "</div>" & vbCrLf
			end if
			show=Replace(user_skin_main,"$show_log$",show_logmore)
			If f_ext = "htm" Or f_ext = "html" Then
				m_commentsmore = Replace(show, "$show_calendar$", "<div id=""calendar""></div><script src='" & 	user_path & "\calendar\" & log_month & ".htm'></script>")
			ElseIf Page="cmd" Then
				m_commentsmore = Replace(show, "$show_calendar$", "<div id=""calendar""></div><" & "%'=Calendar(intYear,intMonth,intDay)%" & ">")
			Else
				m_commentsmore = Replace(show, "$show_calendar$", "<div id=""calendar"">" & oblog.readfile(user_path & "\calendar", log_month & ".htm") & "</div>")
			End If
	end sub

	public sub update_index(resp)
		Dim injsfile
		dim show_topic,show_emot,show_addtime,show_logtext,show_author,show_loginfo,show_more,show_logcyc,show_topictxt,show_logmore,show,rssubject,strtmp,xmlstr,rstmp,start,i,strlogn,strlist,substr
		xmlstr = "<?xml version=""1.0"" encoding=""GB2312""?>" & vbCrLf
		'如果启用了二级或顶级域名，则不使用此项目
		if (true_domain=0 And Cint(oblog.CacheConfig(5))=0)  then
			xmlstr = xmlstr&"<?xml-stylesheet type=""text/xsl"" href="""&blogurl&"oblogstyle/rss.xsl""?>"& vbCrLf
		End If
		xmlstr = xmlstr& "<rss version=""2.0"">"& vbCrLf
		If IsNull(BlogName) Then BlogName=""
		xmlstr = xmlstr & "<channel>" & vbCrLf & "<title><![CDATA[" & Replace(BlogName,"&nbsp;"," ") & "]]></title>" & vbCrLf
		if true_domain=1 then
			xmlstr = xmlstr & "<link>http://" & user_domain & "/index." & f_ext & "</link>" & vbCrLf
		else
			xmlstr = xmlstr & "<link>" & Trim(oblog.CacheConfig(3)) & user_path & "/index." & f_ext & "</link>" & vbCrLf
		end if
		xmlstr = xmlstr & "<description><![CDATA[" & Replace(BlogName,"&nbsp;"," ") & "]]></description>" & vbCrLf
		'处理用户信息
		'update_info user_id
		'处理群组
		'update_mygroups user_id
		if user_showlog_num=0 then user_showlog_num=1
		Set rs = oblog.execute("select subjectid,subjectname from oblog_subject where userid="&user_id)
		While Not rs.EOF
			substr = substr & rs(0) & "!!??((" & rs(1) & "##))=="
			rs.movenext
		Wend
		set rs=oblog.execute("select top "&user_showlog_num&" face,topic,subjectid,logid,istop,addtime,logtext,ishide,commentnum,showword,ispassword,iis,trackbacknum,isbest,blog_password,author,logfile,ishide,authorid,passcheck,Abstract,isneedlogin,viewscores,IsSpecial,viewgroupid from oblog_log where (userid=" & user_id & " or authorid=" & user_id & ")  and isdraft=0 and isdel=0 order by istop desc,addtime desc")
		while not rs.eof
			strtmp=""
			If rs("passcheck") = 0 Then
				strtmp = "此日志需要管理员审核后才可浏览。"
			Else
				If rs("ishide") = 1  Then
					strtmp = "此日志为隐藏日志，仅好友可浏览，<a href='" & blogurl & "more.asp?id=" & rs("logid") & "'>点击进入验证页面</a>。"
				ElseIf rs("ispassword") <> ""Then
					strtmp = "<form method='post' action='" & blogurl & "more.asp?id=" & rs("logid") & "' target='_blank'>请输入日志访问密码：<input type=""password"" size=""15"" name=""password"" />  <input type=""submit"" value=""提交""></form>"
				ElseIf OB_IIF(rs("isneedlogin"),0) = 1  Then
					strtmp = oblog.filt_html_b(rs("Abstract"))	&"<br /><br />此日志需要登录后才可浏览，<a href='" & blogurl & "more.asp?id=" & rs("logid") & "'>点击进入验证页面</a>。"
				ElseIf OB_IIF(rs("viewscores"),0) > 0   Then
					strtmp = oblog.filt_html_b(rs("Abstract"))	&"<br /><br />此日志需要消费积分<strong>"&rs("viewscores")&"</strong>才可浏览，<a href='javascript:void(null);' onclick=""if(confirm('浏览此日志需消费积分"&rs("viewscores")&"，确认浏览？')==true)self.location='" & blogurl & "more.asp?id=" & rs("logid")  & "';return false;"">点击进入验证页面</a>。"
				ElseIf OB_IIF(rs("viewgroupid"),0) > 0   Then
					strtmp = oblog.filt_html_b(rs("Abstract"))	&"<br /><br />此日志需要特定用户组才可浏览，<a href='" & blogurl & "more.asp?id=" & rs("logid") & "'>点击进入验证页面</a>。"
				Else
					If OB_IIF(rs("isspecial"),0) > 0 And Not showpwblog Then
						strtmp = "此日志为特殊日志，<a href='" & blogurl & "more.asp?id=" & rs("logid") & "'>点击进入验证页面</a>"
					End if
				End If
			End if
			'if rs("face")="0" then show_emot="" else	show_emot="<img src=../../images/face/"&rs(0)&".gif >"
			if user_nickname="" or isnull(user_nickname) then
				show_author=user_username
			else
				show_author=user_nickname
			end if
			if rs("authorid")<>user_id then show_author=rs("author")
			show_addtime=rs("addtime")
			If show_addtime="" Or IsNull(show_addtime) Then show_addtime="2007-01-01 0:00:00"
			show_topic=show_emot
			if rs("istop")=1 then show_topic="[置顶]"
			if rs("subjectid")>0 then
					show_topic = show_topic & "<a href=""" & user_truepath&"cmd."&f_ext&"?uid="&user_id&"&do=blogs&id=" & rs("subjectid") & """>[" & oblog.filt_html(getsubname(rs("subjectid"),substr)) & "]</a>"
			end if
			show_topictxt="<a href="""&log_truepath&rs("logfile")&""">"&oblog.filt_html(rs("topic"))&"</a>"
			if rs(13)=1 Then show_topictxt = show_topictxt & "　<img src=" & blogurl & "images/jhinfo.gif >"
			show_topic=show_topic&show_topictxt
			if rs("istop")=1 then show_topictxt="[置顶]"&show_topictxt
			show_loginfo=show_author&" 发表于 "&show_addtime
			show_more="<a href="""&log_truepath&rs("logfile")&""">阅读全文<span id=""ob_logr"&rs("logid")&"""></span></a>"
			show_more=show_more&" | "&"<a href="""&log_truepath&rs("logfile")&"#comment"">回复<span id=""ob_logc"&rs("logid")&"""></span></a> <span id = ""ob_logm"&rs("logid")&"""> </span>"
			show_more = show_more & " | <a href=""" & blogurl & "showtb.asp?id=" & rs(3) & """ target=""_blank"">引用通告<span id=""ob_logt" & rs("logid") & """></span></a>"
			if strtmp<>"" then
				show_logtext=strtmp
			else
				if rs("Abstract")="" or IsNull(rs("Abstract"))  then
					show_logtext=rs(6)
'					show_logtext=oblog.filt_html_b(show_logtext)
'					show_logtext=trimlog(show_logtext,rs(9))
					show_logtext=trimlog(show_logtext,rs("showword"))
					if left(show_logtext,7)="#isubb#" then
						show_logtext=UBBCode(show_logtext,1)
						show_logtext=Replace(show_logtext,CHR(10),"<br />")
						show_logtext = show_logtext & "<br /><a href='"&log_truepath&rs("logfile")&"'>点击显示全文</a>"
					end if
					show_logtext=Replace(show_logtext,"#isubb#","")
					show_logtext=filtimg(show_logtext)
					if oblog.cacheConfig(45)=1 then	show_logtext=profilthtm(show_logtext)
				else
					show_logtext=oblog.filt_html_b(rs("Abstract"))
				end if
			end If
			show_logtext = "<span id=""ob_logd"&rs("logid")&"""></span> "&show_logtext
			'处理RSS
			strlogn=strlogn&"$"&rs("logid")
			If rs("IsSpecial") = 0 OR IsNull("IsSpecial") Then
				xmlstr = xmlstr & "<item>" & vbCrLf & "<title><![CDATA[" & rs("topic") & "]]></title>" & vbCrLf
				if true_domain=1 then
					xmlstr = xmlstr & "<link>" &rs("logfile") & "</link>" & vbCrLf
				else
					xmlstr = xmlstr & "<link>" & Trim(oblog.CacheConfig(3)) & rs("logfile") & "</link>" & vbCrLf
				end if
				xmlstr = xmlstr & "<description><![CDATA[" & oblog.trueurl(rs(6)) & "]]></description>" & vbCrLf
				xmlstr = xmlstr & "<author>" & Replace(show_author,"&nbsp;"," ") & "</author>" & vbCrLf
				xmlstr = xmlstr & "<pubDate>" & show_addtime & "</pubDate>" & vbCrLf & "</item>" & vbCrLf
			end if
			show_logcyc=Replace(user_skin_showlog,"$show_topic$",show_topic)
			show_logcyc=Replace(show_logcyc,"$show_loginfo$",show_loginfo)
			show_logcyc=Replace(show_logcyc,"$show_logtext$",show_logtext)
			show_logcyc=Replace(show_logcyc,"$show_more$",show_more)
			show_logcyc=Replace(show_logcyc,"$show_emot$",show_emot)
			show_logcyc=Replace(show_logcyc,"$show_author$",show_author)
			show_logcyc=Replace(show_logcyc,"$show_addtime$",show_addtime)
			show_logcyc=Replace(show_logcyc,"$show_topictxt$",show_topictxt)
			show_logcyc=Replace(show_logcyc, "$show_blogtag$", "")
			'列表式样
			strlist=strlist&"<li>"&show_topictxt&"　"&show_author&" <span>"&show_addtime&"</span></li>"&vbcrlf
			rs.movenext
			show_logmore=show_logmore&show_logcyc
		wend
		xmlstr=xmlstr& vbcrlf&"</channel>" & vbcrlf&"</rss>"
		set rstmp=oblog.execute("select count(logid) from oblog_log where (userid=" & user_id & " or authorid=" & user_id & ") and passcheck=1 and isdraft=0 and isdel=0")
		G_P_This=1
		start =  CreateStaticPageBar(rstmp(0),user_showlog_num,0)
		rstmp.Close
		Set rstmp=Nothing
		injsfile = "<script src=""" & blogurl & "count.asp?action=logs&id=" & strlogn & """></script>"
		if user_indexlist=1 then show_logmore="<div id=""subject_index""><ul>全部日志"&strlist&"</ul></div>"
		show=Replace(user_skin_main,"$show_log$",show_logmore&start)
		'ATFLAG首页替换掉文摘显示
		show=Replace(show,"$show_blogzhai$","")
		show=filt_inc(show)
		'处理连接
		if showpwblog=false then
			'show=Replace(show,"$show_calendar$","<!-- #include file=""calendar\"&newcalendar(user_path&"/calendar")&".htm"" -->")
			'show=repl_label(show,"",blogname,user_username&","&user_nickname,blogname,user_siteinfo)
			show = repl_label(show, injsfile, BlogName, user_userName & "," & user_nickName, BlogName, user_siteinfo, newcalendar(blogdir&user_path&"/" & "calendar"))
			if ispwblog=false then
				savefile user_path,"/index."&f_ext,show
				savefile user_path,"/rss2.xml",xmlstr
			else
				savefile user_path, "\rss2.xml", "<?xml version=""1.0"" encoding=""GB2312"" ?>  <rss version=""2.0""><channel><title><![CDATA[ 此blog已加密  ]]>   </title></channel></rss>"
				savefile user_path,"/index."&f_ext,"<script language=javascript>window.location.replace('"&blogurl&"pwblog.asp?action=blog&userid="&user_id&"')</script>"
			end if
			if resp=1 then
				Response.Write("<li><a href="&blogurl&user_path&"/index."&f_ext&" target=_blank>点击查看生成的首页</a></li>")
			elseif resp=2 then
				Response.Redirect(blogurl&user_path&"/index."&f_ext)
			ElseIf resp = 3 Then
				Response.Redirect(blogurl&user_path&"/index."&f_ext)
				'Response.Write("<li><a href="&user_path&"/index."&f_ext&" target=_blank>点击查看生成的团队博客首页</a></li>")
			end if
		else
			m_index = show&injsfile
		end if
	end sub

	Public Sub Update_message(resp)
		Dim show_topic, show_emot, show_addtime, show_logtext, show_author, show_loginfo, show_more, show_logcyc, show_topictxt, show_logmore, show
		Dim homepage_str, user_filepath,strPageBar,lngAll
		Set rs = oblog.Execute("select count(messageid) from oblog_message where userid=" & user_id)
		lngAll=rs(0)
		Set rs = oblog.Execute("select top " & user_showlog_num & " message_user,messagetopic,message,addtime,messageid,homepage,isguest,ishide,ubbedit from oblog_message where userid=" & user_id & " and istate= 1 order by messageid desc")
		If Not rs.EOF Then
			While Not rs.EOF
				If IsNull(rs(5)) Then
					homepage_str = "个人主页"
				Else
					If Trim(Replace(rs(5), "http://", "")) = "" Then
						homepage_str = "个人主页"
					Else
						homepage_str = "<a href=""" &blogurl&"go.asp?url=" & oblog.filt_html(rs(5)) & """ target=""_blank"">个人主页</a>"
					End If
				End If
				show_topic = oblog.filt_html(rs(1)) & "<a name='" & rs(4) & "'></a>"
				If rs(6) = 1 Then
					show_author = oblog.filt_html(rs(0)) & "(游客)"
				Else
					show_author = oblog.filt_html(rs(0))
				End If
				show_addtime = rs(3)
				show_topictxt = show_topic
				show_loginfo = show_author & "发表留言于" & show_addtime
				If rs("ubbedit")= 2 Then
					show_logtext = oblog.FilterUbbFlash(filtscript(rs(2)))
				Else
					show_logtext = oblog.Ubb_Comment(rs(2))
				End if
				'show_logtext = oblog.FilterUbbFlash(filtscript(rs(2)))
				show_more = homepage_str & " | <a href='#cmt'>签写留言</a> | <a href='"&blogurl&"user_messages.asp?action=modify&re=true&id=" & rs(4) & "'>回复</a>"
				show_more = show_more & " | <a href=""" & blogurl & "user_messages.asp?action=del&id=" & rs(4) & """  target=""_blank"">删除</a>"
				if rs("ishide")=1 then
					show_topictxt="悄悄话"
					show_topic="悄悄话"
					show_logtext="此留言为悄悄话。"
					show_more=Replace(show_more,"回复","查看")
				end if
				show_logcyc = Replace(user_skin_showlog, "$show_topic$", show_topic)
				show_logcyc = Replace(show_logcyc, "$show_loginfo$", show_loginfo)
				show_logcyc = Replace(show_logcyc, "$show_logtext$", show_logtext)
				show_logcyc = Replace(show_logcyc, "$show_more$", show_more)
				show_logcyc = Replace(show_logcyc, "$show_emot$", "")
				show_logcyc = Replace(show_logcyc, "$show_author$", show_author)
				show_logcyc = Replace(show_logcyc, "$show_addtime$", show_addtime)
				show_logcyc = Replace(show_logcyc, "$show_topictxt$", show_topictxt)
				show_logcyc = Replace(show_logcyc, "$show_blogtag$", "")
				show_logcyc = Replace(show_logcyc, "$show_blogzhai$", "")
				show_logmore = show_logmore & show_logcyc
				rs.movenext
			Wend
			strPageBar = CreateStaticPageBar(lngAll,user_showlog_num,1)
		Else
			show_logmore = "暂无留言"
			strPageBar =""
		End If
		show_logmore = show_logmore &  strPageBar
		Dim strguest, strart, i
		If oblog.cacheConfig(27) Then strguest = "(游客无须输入密码)" Else strguest = ""
		show_logmore = filt_inc(show_logmore)
		show_logmore = show_logmore & vbCrLf & "<div id=""form_comment"">" & vbCrLf
		show_logmore = show_logmore & "	#gg_usercomment#<a name=""cmt""></a><div class=""title"">签写留言：</div>" & vbCrLf
		show_logmore = show_logmore & "	<form action=""" & blogurl & "savemessage.asp?userid=" & user_id & """ method=""post"" name=""commentform"" id=""commentform"" onSubmit=""return Verifycomment()"">" & vbCrLf
		show_logmore = show_logmore & "		<div class=""d1""><label>昵称：<input name=""UserName"" type=""text"" id=""UserName"" size=""20"" maxlength=""20"" value="""" /></label></div>" & vbCrLf
		show_logmore = show_logmore & "		<div class=""d2""><label>密码：<input name=""Password"" type=""password"" id=""Password"" size=""20"" maxlength=""20"" value="""" /> " & strguest & "</label></div>" & vbCrLf
		show_logmore = show_logmore & "		<div class=""d3""><label>主页：<input name=""homepage"" type=""text"" id=""homepage"" size=""42"" maxlength=""50"" value=""http://"" /></label></div>" & vbCrLf
		show_logmore = show_logmore & "		<div class=""d4""><label>标题：<input name=""commenttopic"" type=""text"" id=""commenttopic"" size=""42"" maxlength=""50"" value="""" /></label></div>" & vbCrLf
		show_logmore = show_logmore & "		<div class=""d5"">" & vbCrLf
		show_logmore = show_logmore & "			<input type=""hidden"" name=""edit"" id=""edit"" value="""" />" & vbCrLf
		show_logmore = show_logmore & "			<div id=""oblog_edit""></div>" & vbCrLf
		show_logmore = show_logmore & "		</div>" & vbCrLf
		show_logmore = show_logmore & "		<div class=""d5""><label for=""ishide"">悄悄话：<input name=""ishide"" type=""checkbox"" id=""ishide""  value=""1"" /></label></div>" & vbCrLf
		show_logmore = show_logmore & "		<div class=""d6""><span id=""ob_code""></span><input type=""submit"" value=""&nbsp;提&nbsp;交&nbsp;"" onclick='oblog_edittext.createTextRange().execCommand(""Copy"");'></div>" & vbCrLf
		show_logmore = show_logmore & "	</form>" & vbCrLf
		show_logmore = show_logmore & "</div>" & vbCrLf
		show_logmore = "<h1 class=""message_title"">留言板首页(<a href=""#cmt"">签写留言</a>)</h1>" & vbCrLf & show_logmore
		show = Replace(user_skin_main, "$show_log$", show_logmore)
		if showpwblog=false then
			show = repl_label(show, "", BlogName & "--留言板", user_userName & "," & user_nickName, BlogName, BlogName, newcalendar(blogdir&user_path & "/calendar"))
			if true_domain=1 then
				user_filepath = "http://"&user_domain & "/message." & f_ext
			else
				user_filepath = user_path & "/message." & f_ext
			end if
			If ispwBlog = False Then
				savefile user_path, "\message." & f_ext, show
			Else
				savefile user_path, "\message." & f_ext, "<script language=javascript>window.location.replace('" & blogurl & "pwblog.asp?action=message&userid=" & user_id & "')</script>"
			End If
			If resp = 1 Then
				Response.Write ("<li><a href=" & user_filepath & " target=_blank>点击查看留言板!</a></li>")
			ElseIf resp = 2 Then
				Response.Redirect (user_filepath)
			ElseIf resp = 3 Then
				GoUrl = user_filepath
			End If
		Else
			m_message=show
		End If
	End Sub

	public sub update_info(userid)
		dim show
		show="<ul>"&vbcrlf
		'show=show&"<li><img src=""" & blogdir & oblog.l_uIco&""" widht=""50"" height=""50""/></li>"
		'show=show&"<li>会员昵称:"&OB_IIf(user_nickName,oblog.l_uname)&"</li>"&vbcrlf
		'show=show&"<li>所在城市:"&user_province&user_city&"</li>"&vbcrlf
		'show=show&"<li>会员等级:"&oblog.l_Group(1,0)&"</li>"&vbcrlf
		show=show&"<li><a href="""&user_truepath&"cmd." & f_ext &"?uid="&user_id&"&do=info"">详细信息</a></li><li><a href="""&blogurl&"user_index.asp?url=user_url.asp?action=add$mainuserid="&user_id&"$surl="&user_truepath&"rss2.xml$stitle="&Server.urlencode(blogname)&""" target=""_blank"">站内订阅("&user_subnum&")</a></li></ul>"
		show=show&"<ul><li><a href="""&blogurl&"user_index.asp?url=user_friends.asp?action=add$friendname="&user_username&""" target=""_blank"">加为好友</a></li><li><a href=""javascript:openScript('"&blogurl&"user_pm.asp?action=send&incept="&user_username&"',450,400)"">发送短信</a></li></ul>"
		show=show&"<ul><li>日志:"&user_log_count&"</li>"&vbcrlf
		show=show&"<li>评论:"&user_comment_count&"</li></ul>"&vbcrlf
		show=show&"<ul><li>留言:"&user_message_count&"</li>"&vbcrlf
		show=show&"<li>访问:<span id=""site_count""></span></li></ul>"&vbcrlf
		If OBLOG.CacheConfig(81) = "1" And PlayerType = 0 Then
			If Not IsNull(PassPort_userid) And PassPort_userid >0 Then
				SaveXML "aobomusic","if (chkdiv('aobomusic')) {set_innerHTML('aobomusic','<script language=""Javascript"" src=""http://music.aobo.com/u/"&passport_userid&"/js/?oblog"" charset=""utf-8""></script>')}",True
			End if
		End if
		if showpwblog or showpwlog then m_info=show
		SaveXML "info",show,True
	end sub

	public sub update_placard(userid)
		dim show
		show=filtskinpath(filt_inc(user_placard))
		if showpwblog or showpwlog then m_placard=show
		SaveXML "placard",show,True
	end sub

	public sub update_links(userid)
		dim show
		set rs=oblog.execute("select * from oblog_friendurl where userid="&userid&" order by ordernum asc")
		while not rs.eof
			if rs("urltype")=0 then
				show=show&"<li><a href='"&rs("url")&"' target='_blank'>"&rs("urlname")&"</a></li>"
			else
				show=show&"<li><a href='"&rs("url")&"' target='_blank'><img src='"&rs("logo")&"'></a></li>"
			end if
			rs.movenext
		wend
		show=show&user_links&vbcrlf
		show=filtskinpath(filt_inc(show))
		if showpwblog or showpwlog then m_links=show
		SaveXML "links",show,True
	end sub

	public sub update_newblog(userid)
		dim n,show
		n=Int(user_shownewlog_num)
'		set rs=oblog.execute("select top "&n&" topic,addtime,logfile from [oblog_log] where userid="&userid&" and isdraft=0 and passcheck=1 and isdel=0 order by addtime desc")
		set rs=oblog.execute("select top "&n&" topic,addtime,logfile from [oblog_log] where userid="&userid&" and isdraft=0 and isdel=0 order by addtime desc")
		if not rs.eof then show="<ul>"& vbcrlf
		while not rs.eof
			show=show&"<li><a href="""&log_truepath&rs(2)&""" title=""发表于"&rs(1)&""">"&oblog.filt_html(left(rs(0),18))&"</a></li>"&vbcrlf
			rs.movenext
			if rs.eof then show=show&"</ul>"& vbcrlf
		wend
		if showpwblog or showpwlog then m_newblog=show
		SaveXML "newblog",show,True
	end sub

	Public Sub Update_newmessage(userid)
		Dim n, show, userdir, ustr
		n = CLng(user_shownewmessage_num)
		show = "<ul><li><a href="""&user_truepath&"message." & f_ext & "#cmt""><strong>签写留言</strong></a></li>"
		Set rs = oblog.Execute("select top " & n & " user_dir,messagetopic,b.addtime,message_user,messageid,messagefile from oblog_user a,oblog_message b where b.userid=" & userid & " and a.userid=b.userid and b.istate= 1 AND b.ishide=0 order by messageid desc")
		While Not rs.EOF
			ustr = user_truepath&"message." & f_ext & "#" & rs("messageid")
			show = show & "<li><a href=""" & ustr & """ title=""" & oblog.filt_html(rs("message_user")) & "发表于" & rs("addtime") & """ >" & oblog.filt_html(Left(rs("messagetopic"), 18)) & "</a></li>" & vbCrLf
			rs.movenext
		Wend
		show = show & "</ul>" & vbCrLf
		if showpwblog or showpwlog then m_newmessage=show : exit Sub
		SaveXML "newmessage",show,True
	End Sub

	public sub update_mygroups(userid)
		dim show
		set rs=oblog.execute("select top 6 a.t_name,a.teamid,a.t_ico from oblog_team a,oblog_teamusers  b where a.teamid=b.teamid and a.istate=3 and (b.state=3 or b.state=5 ) and userid ="&userid)
		while not rs.eof
			show=show&"<li><a href=" &blogurl & "group.asp?gid="&rs(1)&" target=""_blank""><img src=""" & ProIco(rs(2),2) & """ class=""group_logo"" /><br />"&oblog.filt_html(left(rs(0),18))&"</a></li>"&vbcrlf
			rs.movenext
		wend
		if showpwblog or showpwlog then
			m_mygroups=show
		Else
			SaveXML "mygroups",show,True
		end if
	end sub

	public sub update_friends(userid)
		dim show
		set rs=oblog.execute("select top 6 a.username,a.nickname,a.user_icon1,a.userid from oblog_user a,oblog_friend  b where a.userid=b.friendid and b.userid ="&userid&" and b.isblack=0 order by b.id desc")
		while not rs.eof
			show=show&"<li><a href=" &blogurl & "go.asp?userid="&rs(3)&" target=_blank><img src=""" & blogurl & OB_IIF(rs(2),"images/ico_default.gif") & """ class=""ob_face"" /><br />"&OB_IIF(rs(1),rs(0))&"</a></li>"&vbcrlf
			rs.movenext
		wend
		if showpwblog or showpwlog then
			m_friends=show
		Else
			SaveXML "myfriend",show,True
		end if
	end sub

	public sub update_comment(userid)
		dim n,show
		n=Int(user_shownewcomment_num)
		set rs=oblog.execute("select top "&n&" oblog_comment.commenttopic,oblog_comment.addtime,oblog_comment.comment_user,oblog_comment.commentid,oblog_log.logfile from oblog_log,oblog_comment where oblog_comment.mainid=oblog_log.logid and oblog_comment.istate= 1 and oblog_comment.userid="&userid&" and oblog_log.isdel=0 and oblog_comment.isdel=0 order by commentid desc")
		if not rs.eof then show="<ul>"& vbcrlf
		while not rs.eof
			show=show&"<li><a href="""&log_truepath&rs("logfile")&"#"&rs("commentid")&""" title="""&oblog.filt_html(rs("comment_user"))&"发表于"&rs("addtime")&""">"&oblog.filt_html(left(rs("commenttopic"),18))&"</a></li>"& vbcrlf
			rs.movenext
			if rs.eof then show=show&"</ul>"& vbcrlf
		wend
		if showpwblog or showpwlog then m_comment=show
		SaveXML "comment",show,True
	end sub

	'生成用户的日志分类
	Public Sub Update_Subject(userid)
		Dim n, show
		show = "<ul>" & vbCrLf & "<li><a href=""" & user_truepath&"index." & f_ext & """ title=""首页"">首页</a>"
		'show = show & "<li><a href=""" & blogdir & "user_index.asp"" target=""blank"">管理</a></li>"
		 show = show & vbCrLf & " <a href="""&user_truepath&"cmd." & f_ext &"?uid="&user_id&"&do=album"" title=""相册"">相册</a> "
		 show = show & vbCrLf & " <a href="""&user_truepath&"cmd."&f_ext&"?uid=" & user_id &"&do=tags"" title=""标签"">标签</a>"
		show=show&"</li>"
		Set rs = oblog.Execute("select Subjectid,SubjectName,Subjectlognum from oBlog_Subject where  userid=" & userid & " and Subjecttype=0 AND (ishide = 0  OR ishide IS NULL) order by ordernum")
		While Not rs.EOF
			show = show & "<li><a href=""" & user_truepath & "cmd."&f_ext&"?do=blogs&id=" & rs("Subjectid") & "&uid="&user_id&""" title=""" & oblog.filt_html(rs("SubjectName")) & """>" & oblog.filt_html(rs("SubjectName")) & "(" & rs("Subjectlognum") & ")" & "</a></li>" & vbCrLf
			rs.movenext
		Wend
		show = show & "</ul>" & vbCrLf
		'show1 = Replace(show, "<div id=""subject"">", "<div id=""subject_l"">")
		if showpwblog or showpwlog then m_subject=show:m_subject_l=show : exit Sub
		SaveXML "subject",show,True
		'使用一个文件，用不同的div id控制格式savefile
	End Sub

	Public Sub Update_calendar(logid)
		Dim c_year, c_year1,c_month, c_day, logdate, today, tomonth, toyear, sql, s, count, b, c
		Dim thismonth, thisdate, thisyear, startspace, NextMonth, NextYear, PreMonth, PreYear, linkTrue
		Dim linkdays, selectdate, linkcount, ccode
		Dim CommondFile
		CommondFile= user_truepath&"cmd."&f_ext&"?uid="&user_id&"&do=month&month="
		ReDim linkdays(2, 0)
		Set rs = oblog.Execute("select addtime from oBlog_log where isdel=0 and oBlog_log.logid=" & Int(logid))
		If rs.EOF Then Exit Sub
		selectdate = rs(0)
		c_year = CInt(Year(selectdate))
		c_month = CInt(Month(selectdate))
		c_day = CInt(Day(selectdate))
		logdate = c_year & "-" & c_month
		If is_sqldata Then
			Dim cmd, rs
			Set cmd = Server.CreateObject("ADODB.Command")
			Set cmd.ActiveConnection = conn
			cmd.CommandText = "ob_calendar"
			cmd.CommandType = 4
			cmd("@logdate") = logdate
			cmd("@userid") = user_id
			Set rs = cmd.Execute
			Set cmd = Nothing
		Else
			sql = "select addtime,logfile from oBlog_log WHERE year(addtime)=" & c_year & " and month(addtime)=" & c_month & " and isdel=0  and userid=" & user_id & " ORDER BY addtime DESC "
			Set rs = oblog.Execute(sql)
		End If
		Dim theday
		theday = 0

		Do While Not rs.EOF
			If Day(rs("addtime")) <> theday Then
				theday = Day(rs("addtime"))
				ReDim Preserve linkdays(2, linkcount)
				linkdays(0, linkcount) = Month(rs("addtime"))
				linkdays(1, linkcount) = Day(rs("addtime"))
				'linkdays(2, linkcount) = blogdir & rs("logfile")
				linkdays(2, linkcount)=user_truepath&"cmd."&f_ext&"?uid="&user_id&"&do=day&day=" & CStr(CDate(Year(rs("addtime")) & "-" & Month(rs("addtime")) & "-" & Day(rs("addtime"))))
				linkcount = linkcount + 1
			End If
			rs.movenext
		Loop
		Set rs = Nothing
		Dim mdays(12)
		mdays(0) = ""
		mdays(1) = 31
		mdays(2) = 28
		mdays(3) = 31
		mdays(4) = 30
		mdays(5) = 31
		mdays(6) = 30
		mdays(7) = 31
		mdays(8) = 31
		mdays(9) = 30
		mdays(10) = 31
		mdays(11) = 30
		mdays(12) = 31
		'今天的年月日
		today = Day(oblog.ServerDate(Now()))
		tomonth = Month(oblog.ServerDate(Now()))
		toyear = Year(oblog.ServerDate(Now()))
		'指定的年月日及星期
		thismonth = c_month
		thisdate = c_day
		thisyear = c_year
		If IsDate("February 29, " & thisyear) Then mdays(2) = 29
		'确定日历1号的星期
		startspace = Weekday(thismonth & "-1-" & thisyear) - 1
		NextMonth = c_month + 1
		NextYear = c_year+1
		If NextMonth > 12 Then
			NextMonth = 1
			NextYear = NextYear + 1
		End If
		PreMonth = c_month - 1
		'PreYear = c_year-1
		PreYear = c_year
		If PreMonth < 1 Then
			PreMonth = 12
			PreYear = PreYear - 1
		End If
		ccode = "<table width='100%' class=""year_"&thisyear&" month_"&thismonth&""">" & vbCrLf
		ccode = ccode & "<thead>" & vbCrLf
		'ccode = ccode & "<caption>" & mName(thismonth) & thisyear & "</caption><tr>" & vbCrLf
		ccode = ccode & "<caption><a href="""& CommondFile & ((c_year-1) & Right("0" & c_month,2)) &""" title=""上一年""><span class=""arrow"">&lt;&lt;</span></a>&nbsp;&nbsp;<a href=""" & CommondFile & PreYear& Right("0" & preMonth,2)&""" title=""上一月""><span class=""arrow"">&lt;</span></a>&nbsp;"& toyear &"<a href=""" & CommondFile & Year(oblog.ServerDate(Date)) & Right("0" & Month(oblog.ServerDate(Date)),2) & """ title=""返回当月""> - </a>"& c_month&"&nbsp;<a href="""& CommondFile & c_year& Right("0" & NextMonth,2) &""" title=""下一月""><span class=""arrow"">&gt;</span></a>&nbsp;&nbsp;<a href=""" & CommondFile & NextYear & Right("0" & c_month,2) &""" title=""下一年""><span class=""arrow"">&gt;&gt;</span></a></caption>" & vbCrLf
		'ccode = ccode & "<caption><a href=""" & CommondFile & c_year& Right("0" & preMonth,2)&""" title=""上一月""><span class=""arrow""><</span></a> "& toyear &"  <a href=""" & CommondFile & Year(oblog.ServerDate(Date)) & Right("0" & Month(oblog.ServerDate(Date)),2) & """ title=""返回当月"">-</a> "& c_month&" <a href="""& CommondFile & c_year& Right("0" & NextMonth,2) &""" title=""下一月""><span class=""arrow"">></span></a></caption>" & vbCrLf
		ccode = ccode & "<tr class=""week"">" & vbCrLf
		ccode = ccode & "<th class=""sun"">日</th>" & vbCrLf
		ccode = ccode & "<th class=""mon"">一</th>" & vbCrLf
		ccode = ccode & "<th class=""Tue"">二</th>" & vbCrLf
		ccode = ccode & "<th class=""Wen"">三</th>" & vbCrLf
		ccode = ccode & "<th class=""Thu"">四</th>" & vbCrLf
		ccode = ccode & "<th class=""Fri"">五</th>" & vbCrLf
		ccode = ccode & "<th class=""Sat"">六</th>" & vbCrLf
		ccode = ccode & "</tr>" & vbCrLf
		ccode = ccode & "</thead>" & vbCrLf
		ccode = ccode & "<tbody>" & vbCrLf
		ccode = ccode & "<tr>" & vbCrLf
		For s = 0 To startspace - 1
			ccode = ccode & "<td align=""center""></td>" & vbCrLf
		Next
		count = 1
		While count <= mdays(thismonth)
			For b = startspace To 6
				ccode = ccode & "<td align=""center"""
				If count=thisdate+1 Then
					ccode = ccode & " class=""today"" title=""今天"""
				End if
				ccode = ccode & ">"
				linkTrue = "False"
				For c = 0 To UBound(linkdays, 2)
					If linkdays(0, c) <> "" Then
						If linkdays(0, c) = thismonth And linkdays(1, c) = count Then
							ccode = ccode & "<a href=""" & linkdays(2, c) & """ title=""查看"&thisyear&"年"&thismonth&"月"&count&"日的日志"">"
							linkTrue = "True"
						End If
					End If
				Next
				If count <= mdays(thismonth) Then ccode = ccode & count
				If linkTrue = "True" Then ccode = ccode & "</a>"
				ccode = ccode & "</td>" & vbCrLf
				count = count + 1
			Next
			If count > mdays(thismonth) Then
				ccode = ccode & "</tr>" & vbCrLf
			  Else
				ccode = ccode & "</tr><tr>" & vbCrLf
			End If
			startspace = 0
		Wend
		ccode = ccode & "</tbody>" & vbCrLf
		ccode = ccode & "</table>" & vbCrLf
		If Int(c_month) < 10 Then c_month = "0" & c_month
		'ccode = "<div id=""calendar"">" & ccode & "</div>"
		savefile user_path, "\calendar\" & c_year & c_month & ".htm", ccode
	End Sub

	Public Function filt_pwblog(show, log_title)
		update_info (user_id)
		update_subject (user_id)
		update_newblog (user_id)
		update_newmessage (user_id)
		update_links (user_id)
		update_comment (user_id)
		Update_placard (user_id)
		Update_friends (user_id)
		Update_mygroups (user_id)
		'show=Replace(show,"$show_calendar$",oblog.readfile(user_path&"\calendar",log_month&".htm"))
		show = Replace(show, "$show_userid$",user_id)
		show = Replace(show, "$show_placard$", "<div id=""placard"">"&m_placard&"</div>")
		show = Replace(show, "$show_subject$", "<div id=""subject"">"&m_subject&"</div>")
		show = Replace(show, "$show_subject_l$", "<div id=""subject_l"">"&m_subject&"</div>")
		show = Replace(show, "$show_newblog$", "<div id=""newblog"">"&m_newblog&"</div>")
		show = Replace(show, "$show_comment$", "<div id=""comment"">"&m_comment&"</div>")
		show = Replace(show, "$show_newmessage$", "<div id=""newmessage"">"&m_newmessage&"</div>")
		show = Replace(show, "$show_links$", "<div id=""links"">"&m_links&"</div><div id=""gg_userlinks""></div>")
		If InStr(show,"$show_music$") Then
		show = Replace(show, "$show_info$", "<div id=""info"">"&m_info&"</div>")
		show = Replace(show, "$show_music$", "<div id=""aobomusic""></div>")
		Else
		show = Replace(show, "$show_info$", "<div id=""info"">"&m_info&"</div><div id=""aobomusic""></div>")
		End If
		show = Replace(show, "$show_blogname$", blogname)
		show = Replace(show, "$show_myfriend$", "<div id=""myfriend"">" &m_friends&"</div>")
		show = Replace(show, "$show_mygroups$", "<div id=""mygroups"">"&m_mygroups&"</div>")
		show = Replace(show, "$show_blogurl$", user_trueurl)
		'show = Replace(show, "$show_photo$", "<div id=""ob_miniphoto""></div><script>var so = new SWFObject("""&blogurl&"miniphoto.swf?blogurl="&blogurl&"&userid="&user_id&"&gourl="&user_truepath&"cmd."&f_ext&"?uid="&user_id&"$do=album"", ""miniphoto"", ""100%"", ""180"", ""9"", ""#FFFFFF"");so.addParam(""wmode"", ""transparent"");so.write(""ob_miniphoto"");<script>")
		show = Replace(show, "$show_photo$", "<div id=""ob_miniphoto""><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0' width='100%' height='180' align='middle'><param name=""wmode"" value=""transparent"" /><param name='movie' value='"&blogurl&"miniphoto.swf?blogurl="&blogurl&"&userid="&user_id&"&gourl="&user_truepath&"cmd."&f_ext&"?uid="&user_id&"$do=album' /><param name='quality' value='high' /><embed src='"&blogurl&"miniphoto.swf?blogurl="&blogurl&"&userid="&user_id&"&gourl="&user_truepath&"cmd."&f_ext&"?uid="&user_id&"$do=album' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='100%' height='180'></embed></object></div>")
		show = Replace(show, "#gg_usercomment#", "<div id=""gg_usercomment""></div>")
		show=Replace(show,"$show_xml$","<div id=""xml""><span id=""txml""></span><br /><a href="""&user_truepath&"rss2.xml"" target=""_blank""><img src='" & blogurl & "images/xml.gif' width='36' height='14' border='0' /></a></div>")
		show=Replace(show,"$show_search$","<div id=""search""></div>" )
		show=Replace(show,"$show_login$","<div id=""ob_login""></div>")
		show="<script src=""inc/main.js"" type=""text/javascript""></script>"&VbCrLf&show
		show="<link href=""OblogStyle/OblogUserDefault4.css"" rel=""stylesheet"" type=""text/css"" />"&VbCrLf&"</head>"&VbCrLf&"<body><span id=""gg_usertop""></span>"&show
		show=show&"<div id=""powered""><a href=""http://www.oblog.cn"" target=""_blank""><img src=""images/oblog_powered.gif"" border=""0"" alt=""Powered by Oblog."" /></a></div>"&VbCrLf&"<span id=""gg_userbot""></span></body>"&VbCrLf&"</html>"
		show="<title>"&log_title&"</title>"&VbCrLf&show
		show="<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"&VbCrLf&show
		show="<meta http-equiv=""Content-Language"" content=""zh-CN"" />"&VbCrLf&show
		show="<html>"&VbCrLf&"<head>"&VbCrLf&show
		If InStr(show, "<div id=""oblog_edit"">") Then
			show = show & "<script src=""" & blogurl & "count.asp?action=code31""></script>" & vbCrlf
			show = show & "<script src=""" & blogurl & "commentedit.asp""></script>" & vbCrlf
		End If
		If InStr(show, "<div id=""blogzhai"">") Then
			show = show & "<script src=""" & blogurl & "inc/inc_zhai.js""></script>" & vbCrlf
		End If
		show =repl_JS(show)
		show = Replace(show,"#CreateFunctionPage#","")
		filt_pwblog = show
	End Function

	Public Sub remove_user_skin_cache()
		If Application(Cache_Name & "_user_skin_main") <> "" Then
				Application.Lock
				Application.Contents.Remove (Cache_Name & "_user_skin_main")
				Application.unLock
		End If
		If Application(Cache_Name & "_user_skin_showlog") <> "" Then
				Application.Lock
				Application.Contents.Remove (Cache_Name & "_user_skin_showlog")
				Application.unLock
		End If
	End Sub

	public sub update_alllog_admin(uid)
		dim rs1,i,p
        uid = CLng(uid)
        oblog.CreateUserDir uid, 0
        userid = uid
		Set rs1 = Server.CreateObject("Adodb.RecordSet")
        rs1.open "select logid from oBlog_log where userid=" & uid & " and isdraft=0 ORDER BY logid", conn, 1, 1
        While Not rs1.EOF
            p = rs1.recordcount + 1
            progress Int(i / p * 100), "更新ID为" & rs1(0) & "的日志..."
            Update_log rs1(0), 0
            Update_calendar rs1(0)
            i = i + 1
            rs1.movenext
        Wend
        rs1.Close
		set rs1=nothing
		update_usite(uid)
	end sub

    Public Sub Update_alllog(uid,lastlogid)
        uid = CLng(uid)
        oblog.CreateUserDir uid, 0
        userid = uid
        update_partlog uid,lastlogid
        update_usite(uid)
    End Sub

    Public Sub Update_subjectlog(uid,subjectid)
        uid = CLng(uid)
        userid = uid
		Dim p,i,rs1
		set rs1=Server.CreateObject("Adodb.RecordSet")
		rs1.open "select logid FROM oblog_log WHERE userid="&uid&" and isdraft=0 and isdel=0 and subjectid = "&subjectid & " ORDER BY logid ",conn,1,1
		if rs1.eof then
			rs1.Close
			set rs1=nothing
			progress 100, "更新当前专题日志完成!"
			exit sub
		end if
		While Not rs1.eof
			p=rs1.recordcount+1
			progress Int(i/p*100),"更新ID为"&rs1(0)&"的日志..."
			update_log rs1(0),0
			update_calendar rs1(0)
			rs1.movenext
		Wend
		rs1.close
		set rs1=Nothing
        update_usite(uid)
    End Sub

	public sub update_partlog(uid,lid)
		dim p,i,rs1,lastid
		uid=Int(uid)
		lid=CLng(lid)
		i=1
		userid=uid
		set rs1=Server.CreateObject("Adodb.RecordSet")
		rs1.open "select TOP " & P_BLOG_UPDATEPAUSE &" logid FROM oblog_log WHERE userid="&uid&" and isdraft=0 and isdel=0 and logid>"&lid & " ORDER BY logid ",conn,1,1
		if rs1.eof then
			rs1.Close
			set rs1=nothing
			progress 100, "更新所有日志完成!"
			exit sub
		end if
		While Not rs1.eof
			p=rs1.recordcount+1
			progress Int(i/p*100),"更新ID为"&rs1(0)&"的日志..."
			update_log rs1(0),0
			update_calendar rs1(0)
			lastid=rs1(0)
			i=i+1
			rs1.movenext
		Wend
		rs1.close
		set rs1=oblog.execute("select top 1 logid from oblog_log where userid=" & uid & " and isdraft=0 and isdel=0 and logid>"&lastid)
		if rs1.eof then
			set rs1=Nothing
			progress 100, "更新所有日志完成!"
			exit Sub
		End if
		rs1.Close
		set rs1=Nothing
		Dim ttime
		ttime = oblog.CacheConfig(28)
		If ttime <> "" Then ttime = CLng (ttime) Else ttime = 5
		If ttime > 60 Then ttime = 5
		If ttime > 0 Then
			progress 100, ttime&"秒后自动更新后面的日志，请不要刷新页面..."
			with Response
				.Write "<script language=JavaScript>var progress=document.getElementById(""progress"");var secs = "&ttime&";var wait = secs * 1000;"
				.write "for(i = 1; i <= secs; i++){window.setTimeout(""Update("" + i + "")"", i * 1000);}"
				.write "function Update(num){if(num != secs){printnr = (wait / 1000) - num;document.getElementById(""pstr"").innerHTML=printnr+""秒后自动更新后面的日志，请不要刷新页面..."";progress.style.width=(num/secs)*100+""%"";progress.innerHTML=""剩余""+printnr+""秒""}}"
				.write "setTimeout(""window.location='user_update.asp?action=update_alllog&lastlogid="&lastid&"'"","&Int(ttime*1000)&");</script>"
			end with
			Response.Flush()
			Response.End()
        End If
		set rs1=nothing
	end Sub

	public sub update_usite(uid)
		dim p
		p = 14
		progress Int(1/p*100),"更新首页..."
		update_index 0
		progress Int(2/p*100),"更新站点信息文件..."
		update_info uid
		progress Int(3/p*100),"生成新日志列表文件..."
		update_newblog(uid)
		progress Int(4/p*100),"更新最新留言..."
		update_newmessage uid
		progress Int(5/p*100),"生成首页日志分类文件..."
		update_subject(uid)
		progress Int(8/p*100),"更新留言板..."
		update_message 0
		progress Int(9/p*100),"更新最新回复..."
		update_mygroups uid
		progress Int(10/p*100),"更新" &oblog.CacheConfig(69)& "列表..."
		update_comment uid
		progress Int(11/p*100),"更新公告..."
		update_placard uid
		progress Int(12/p*100),"更新友情连接..."
		update_links uid
		progress Int(13/p*100),"更新blogname..."
		update_friends uid
		update_blogname
		CreateFunctionPage
		progress Int(14/p*100),"重新发布完成!"
	end sub

	public sub update_blogname()

		SaveXML "blogname",oblog.filt_html(blogname),True
	end sub
	Public Sub savefile(dirstr, fname, str)
		On Error Resume Next
		Dim dirstr1, divid
		if dirstr="" then
			Response.Write("用户目录不能为空！")
			Response.end
		end if
		dirstr1 = Server.Mappath(blogdir&dirstr)
		'Response.Write "目录2：" & dirstr1  & "<br>"
		'以下转为js格式
		If (Left(fname, 5) = "\inc\" Or Left(fname, 10) = "\calendar\") And (f_ext = "htm" Or f_ext = "html") Then
			If Left(fname, 10) = "\calendar\" Then
				divid = "calendar"
			Else
				divid = Replace(Replace(Replace(fname, "\inc\", ""), ".htm", ""), "show_", "")
			End If
			str = oblog.htm2js_div(str, divid)
		End If
		'以下兼容asp格式,转换路径
		if f_ext="asp" and true_domain=0 then
			if oblog.CacheConfig(57)="0"  then
				str=Replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""")
			else
				if instr(fname,"index.asp") or instr(fname,"message.asp") or instr(fname,"cmd.asp") then
					str=Replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""")
				else
					str=Replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""../../")
				end if
			end if

			if oblog.CacheConfig(57)="0"  then
				str=Replace(str,"<!-- #include file="""&blogdir,"<!-- #include file=""../../")
			else
				if instr(fname,"index.asp") or instr(fname,"message.asp") or instr(fname,"cmd.asp") then
					str=Replace(str,"<!-- #include file="""&blogdir,"<!-- #include file=""../../")
				else
					str=Replace(str,"<!-- #include file="""&blogdir,"<!-- #include file=""../../../../")
				end if
			end if
		end if
		if true_domain=1 then
			if oblog.CacheConfig(57)="0"  then
				str=Replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""")
			else
				if instr(fname,"index."&f_ext) or instr(fname,"message."&f_ext) or instr(fname,"cmd."&f_ext) then
					str=Replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""")
				else
					str=Replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""../../")
				end if
			end if
		end If
		str = Replace(str,"#CreateFunctionPage#","")
		If str = "" Or IsNull(str) Then str = " "
		If objFSO.FolderExists(dirstr1) = False Then objFSO.CreateFolder (dirstr1)
		Call oblog.BuildFile(dirstr1 & Trim(fname), str )
	End Sub

	public Function newcalendar(folderspec)
		On Error Resume Next
		Dim f, f1, fc, nname
		'Response.Write folderspec
		if objFSO.FolderExists(Server.MapPath(folderspec)) then
			Set f = objFSO.GetFolder(Server.MapPath(folderspec))
			Set fc = f.Files
			 nname=0
			For Each f1 in fc
				If IsNumeric(Replace(f1.name,".htm","")) Then
			   if nname<Int(Replace(f1.name,".htm","")) then nname=Int(Replace(f1.name,".htm",""))
			   End If
			Next
			newcalendar = nname
		else
			newcalendar="0"
		end if
	End Function

	'str改为了代入的js包含文件
	Public Function repl_label(show, str, title, author, keyword, desc, calendar)
		On Error Resume Next
		show = Replace(show, "$show_userid$",user_id)
		show = Replace(show, "#gg_usercomment#", "<div id=""gg_usercomment""></div>")
		show = Replace(show, "$show_placard$", "<div id=""placard"">"&oblog.cacheConfig(41)&"</div>")
		If calendar <>"" Then
			show = Replace(show, "$show_calendar$", "<div id=""calendar""><!-- #include file="""&user_truepath&"calendar/" & calendar & ".htm"" --></div>")
		Else
			show = Replace(show, "$show_calendar$", "<div id=""calendar""></div>")
		End if
		show = Replace(show, "$show_xml$", "<div id=""xml""><span id=""txml""></span><br /><br /><a href="""&user_truepath&"rss2.xml"" target=""_blank""><img src='" & blogurl & "images/xml.gif' width='36' height='14' border='0' /></a></div>")
		show = Replace(show, "$show_subject$", "<div id=""subject"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_subject_l$", "<div id=""subject_l"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_newblog$", "<div id=""newblog"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_comment$", "<div id=""comment"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_mygroups$", "<div id=""mygroups"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_myfriend$", "<div id=""myfriend"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_blogurl$", "<span id=""blogurl"">"&user_trueurl&"</span>")
		show = Replace(show, "$show_blogname$", "<span id=""blogname"">"&oblog.cacheConfig(41)&"</span>")
		show = Replace(show, "$show_newmessage$", "<div id=""newmessage"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_links$", "<div id=""links""></div><div id=""gg_userlinks""></div>")
		If InStr(show,"$show_music$") Then
		show = Replace(show, "$show_info$", "<div id=""info"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_music$", "<div id=""aobomusic""></div>")
		Else
		show = Replace(show, "$show_info$", "<div id=""info"">"&oblog.cacheConfig(41)&"</div><div id=""aobomusic""></div>")
		End If
		show = Replace(show, "$show_search$", "<div id=""search"">"&oblog.cacheConfig(41)&"</div>")
		show = Replace(show, "$show_login$", "<div id=""ob_login"">"&oblog.cacheConfig(41)&"</div>")
		'show = Replace(show, "$show_photo$", "<div id=""ob_miniphoto""></div><script>var so = new SWFObject("""&blogurl&"miniphoto.swf?blogurl="&blogurl&"&userid="&user_id&"&gourl="&user_truepath&"cmd."&f_ext&"?uid="&user_id&"$do=album"", ""miniphoto"", ""100%"", ""180"", ""9"", ""#FFFFFF"");so.addParam(""wmode"", ""transparent"");so.write(""ob_miniphoto"");<script>")
		show = Replace(show, "$show_photo$", "<div id=""ob_miniphoto""><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0' width='100%' height='180' align='middle'><param name=""wmode"" value=""transparent"" /><param name='movie' value='"&blogurl&"miniphoto.swf?blogurl="&blogurl&"&userid="&user_id&"&gourl="&user_truepath&"cmd."&f_ext&"?uid="&user_id&"$do=album' /><param name='quality' value='high' /><embed src='"&blogurl&"miniphoto.swf?blogurl="&blogurl&"&userid="&user_id&"&gourl="&user_truepath&"cmd."&f_ext&"?uid="&user_id&"$do=album' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='100%' height='180'></embed></object></div>")

		show = "<link rel=""alternate"" href="""&user_truepath&"rss2.xml"" type=""application/rss+xml"" title=""RSS"" />"&vbCrLf&"<link href=""" & blogurl & "OblogStyle/OblogUserDefault4.css"" rel=""stylesheet"" type=""text/css"" />" & vbCrLf & "<script src=""" & blogurl & "inc/main.js"" type=""text/javascript""></script>" & vbCrLf & "</head>" & vbCrLf & "<body>" & vbCrLf & "<span id=""gg_usertop""></span>" & show
		show = show & "<span id=""gg_userbot""></span>"
		show = show & "<div id=""powered""><a href=""http://www.oblog.cn"" target=""_blank""><img src="""&blogurl&"images/oblog_powered.gif"" border=""0"" alt=""Powered by Oblog."" /></a></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
		Dim showTitle,showFrame
		showTitle = "<title>" & Replace(Replace(Replace(title, "&lt;", "＜"), "&gt;", "＞"), "&nbsp;", "　") & "</title>" & vbCrLf
		showTitle = "<meta name=""description"" content=""" & oblog.filt_html(desc) & """ />" & vbCrLf &showTitle
		showTitle = "<meta name=""keyword"" content=""" & oblog.filt_html(keyword) & """ />" & vbCrLf & showTitle
		showTitle = "<meta name=""author"" content=""" & oblog.filt_html(author) & """ />" & vbCrLf & showTitle
		showTitle = "<meta name=""generator"" content=""oblog"" />" & vbCrLf & showTitle
		showTitle = "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbCrLf & showTitle
		showTitle = "<meta http-equiv=""Content-Language"" content=""zh-CN"" />" & vbCrLf & showTitle
		showTitle = "<html>" & vbCrLf & "<head>" & vbCrLf & showTitle
		'showTitle = "<!--Last Update Time : "&Now()&"-->"&vbCrLf&showTitle
		showFrame = "<frameset rows=""*,45"" frameborder=""no"" border=""0"" framespacing=""0"">" & vbCrLf
		showFrame = showFrame &"<frame src=""index."&F_EXT&""" name=""mainFrame"" id=""mainFrame"" />" & vbCrLf
		showFrame = showFrame & "<frame src=""http://music.aobo.com/u/"&PassPort_userid&"/js/?oblog"" name=""AOBOMusic"" scrolling=""No"" noresize=""noresize"" id=""AOBOMusic"" />" & vbCrLf
		showFrame = showFrame &"</frameset>" & vbCrLf
		showFrame = showFrame &"<noframes><body>" & vbCrLf
		showFrame = showFrame &"</body>" & vbCrLf
		showFrame = showFrame &"</noframes>" & vbCrLf
		showFrame = showFrame &"</html>"
		showFrame = showTitle & showFrame
		If OBLOG.CacheConfig(81) = "1" And PlayerType = 1 Then
			If Not IsNull(PassPort_userid) And PassPort_userid>0 Then
				oblog.BuildFile Server.Mappath(blogdir&user_path &"/default."&f_ext),showFrame
				showTitle = showTitle & "<script>"& vbCrLf
				showTitle = showTitle & "if (top.location == self.location &&self.location.href.lastIndexOf('index."&f_ext&"')>0) { "& vbCrLf
				If CBool(true_domain) Then
				showTitle = showTitle & "	top.location ='"&user_truepath&"default."&f_ext&"';"& vbCrLf
				Else
				showTitle = showTitle & "	top.location ='"&blogdir&user_path&"/default."&f_ext&"';"& vbCrLf
				End If
				showTitle = showTitle & "} "& vbCrLf
				showTitle = showTitle & "</script>"& vbCrLf
			End if
		End if
		show =  showTitle & show
		'html文件，将包含文件改为js包含
		If f_ext = "htm" Or f_ext = "html" Then
			dim jspath
			jspath=blogurl&user_path&"/"
			show = filt_include(show)
			show = show & "<script src=""" & jspath&"calendar/" & calendar & ".htm""></script>" & vbCrlf
		End If
		'以下为shtml，asp，html共用js包含
		If InStr(show, "<div id=""oblog_edit"">") Then
			show = show & "<script src=""" & blogurl & "count.asp?action=code31""></script>" & vbCrlf
			show = show & "<script src=""" & blogurl & "commentedit.asp""></script>" & vbCrlf
		End If
		If InStr(show, "<div id=""blogzhai"">") Then
			show = show & "<script src=""" & blogurl & "inc/inc_zhai.js""></script>" & vbCrlf
		End If
		show = show & str
		show=repl_JS(show)
		repl_label=filtskinpath(show)
	End Function

	'页面通用JS
	public Function repl_JS(str)
		Dim show
		show = str
		show = show & "#CreateFunctionPage#"& vbCrlf
		show = show & "<script src=""" & blogurl & "login.asp?action=showindexlogin""></script>" & vbCrlf
		'XML的载入必须放在页面最底部，否则可能引起以innerHTML插入异常，切记！！！！
		show = show & "<script src=""" & blogurl&"ShowXml.asp?user_group="&user_group&"&user_path="&user_path&"&userid="&user_id&"&blogname="&BlogName&"""></script>" & vbCrlf
		show = show & "<script src=""" & blogurl & "count.asp?action=site&id=" & user_id & """></script>"
		repl_JS=show
	end function

	public sub progress_init1()
		Response.Write("<table width=""400"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">")
		Response.Write("<tr><td style=""text-align:center;""><span id=txt1 name=txt1 ></span></td></tr>")
		Response.Write("<tr><td><img src=""images/bar.gif"" width=0 height=10 id=img1 name=img1 align=absmiddle></td></tr>")
		Response.write("<tr><td><span id=txt2 name=txt2 ></span></td></tr></table>")
	end sub

	public sub progress_init()
		Response.Write("<div class=""progress1""><div class=""progress2"" id=""progress""></div></div><span id=""pstr""></span>")
	end sub

	public sub progress(num,str)
		Response.Write "<script>var progress=document.getElementById(""progress"");progress.style.width ="""&num&"%"";progress.innerHTML="""&num&"%"";document.getElementById(""pstr"").innerHTML="""&str&""";</script>" & VbCrLf
		Response.Flush
	end sub


	'创建一个主功能界面，其功能标签就是$show_log$
	'生成时机
	'发布/修改文章/发表回复/发表留言/模板更新/整站更新
	Public Sub CreateFunctionPage()
		Page="cmd"
		On Error Resume Next
		Dim strPage,str
		str=user_skin_main
		strPage = repl_label(str, "", BlogName, user_userName & "," & user_nickName, "", "", "")
		strPage=Replace(strPage,"$show_userid$",user_id)
		strPage=Replace(strPage,"$show_log$","<div id=oblog_usercontent>"&oblog.CacheConfig(41)&"</div>")
		If CBool(Islightbox) Then
		strPage=Replace(strPage,"<div id=oblog_usercontent>","<link rel=""stylesheet"" href="""&blogurl&"Plus/lightbox2.03.3/lightbox_oblog.asp?f=lightbox.css"" type=""text/css"" media=""screen"" /><script type=""text/javascript"" src="""&blogurl&"Plus/lightbox2.03.3/prototype.js""></script><script type=""text/javascript"" src="""&blogurl&"Plus/lightbox2.03.3/scriptaculous.js?load=effects""></script><script type=""text/javascript"" src="""&blogurl&"Plus/lightbox2.03.3/lightbox_oblog.asp?f=lightbox.js""></script><div id=oblog_usercontent>")
		End If
		strPage = Replace(strPage ,"#CreateFunctionPage#","<script>document.write('<script language=""javascript"" src="""&blogurl&"pagecmd.asp?'+getpara()+'""><\/script>');</script>"& vbCrlf &"<script src="""&blogurl&"commentedit.asp""></script>"& vbCrlf & "<script src="""&blogurl&"count.asp?action=code31""></script>")
		'CMD页面需要评论功能,编辑器、验证码的赋值需放最后
		if f_ext="shtml" or f_ext="asp" then
			Dim objRegExp
			Set objRegExp = New Regexp
			objRegExp.IgnoreCase = True
			objRegExp.Global = True
			objRegExp.Pattern = "<div id=""calendar"">.*?</div>"
			strPage = objRegExp.replace(strPage, "<div id=""calendar"">"&oblog.CacheConfig(41)&"</div>")
			Set objRegExp = Nothing
		end if
		Savefile user_path, "\cmd."&f_ext, strPage
		strPage=""
		Page=""
	End Sub
	'创建静态首页页面的分页导航条
	Private Function CreateStaticPageBar(byval lngAll,byval intPerPage,byval intType)
		Dim strPageBar,strUnit
		G_P_AllRecords=lngAll
		G_P_PerMax=intPerPage
		G_P_FileName=user_truepath&"cmd."&f_ext&"?uid=" & user_id& "&do="
		select Case intType
			Case 0
				G_P_FileName=G_P_FileName & "index"
				strUnit="篇日志"
			Case 1
				G_P_FileName=G_P_FileName & "message"
				strUnit="篇留言"
		End select
		G_P_This=1
		CreateStaticPageBar=oBlog.ShowPage(false, true, strUnit)
	End Function

	Private Function SaveXML(DIVID,ByVal ElementText,IsCDATA)
	    On Error Resume Next
		Dim xmlDoc,userpath
		Set xmlDoc = New Cls_XmlDoc
		'xmlDoc.Unicode = False
		userpath = blogdir&user_path&"/user.config"
		If xmlDoc.LoadXml (userpath) Then
			If DIVID = "aobomusic" Then
				xmlDoc.UpdateNodeText DIVID,ElementText,IsCDATA
			Else
				xmlDoc.UpdateNodeText DIVID,oblog.htm2js_div(ElementText,DIVID),IsCDATA
			End If
		Else
			If xmlDoc.LoadXml (blogdir&"XmlData/user.config") Then
				xmlDoc.SaveAs userpath
			Else
				Response.Write (blogdir&"XmlData/user.config 不存在，无法继续操作！")
				Set XmlDoc = Nothing
				Response.End
			End If
			'递归
			SaveXML DIVID,ElementText,IsCDATA
		End If
		xmlDoc.Save
		Set xmlDoc = Nothing
	End Function
	'删除关联文件
	'logids可能为多个日志，如1,2,3
	Public Sub DeleteFiles(ByVal logids,ByVal userid)
		On Error Resume Next
	    Dim rst, fs, fsize, uid, imgsrc, fid,logid,aLogid,z
		If logids = "" Then
			If userid = "" Then Exit Sub
			Set rst = oblog.Execute ("SELECT logid FROM oblog_log WHERE userid="&Int(userid))
			If Not rst.EOF Then
				While Not rst.EOF
					logids = logids &","&RST(0)
					rst.MoveNext
				Wend
			Else
				Set RST = Nothing
				Exit Sub
			End If
		End If
	    logids=FilterIds(logids)
	    aLogid=Split(logids,",")
	    For z=0 To Ubound(aLogid)
		    logid=aLogid(z)
			'删除DIGG
			Dim RSDIGG
			Set RSDIGG = oblog.Execute ("SELECT COUNT(did),authorid FROM oblog_digg WHERE diggtype = -1 AND logid = " & logid &" GROUP BY authorid ")
			If Not RSDIGG.Eof Then
				oblog.GiveScore "",-1*Abs(oblog.CacheScores(22))*RSDIGG(0),RSDIGG(1)
			End If
			oblog.Execute ("DELETE FROM oblog_userdigg WHERE logid = "&logid)
			oblog.Execute ("DELETE FROM oblog_digg WHERE logid = "&logid)
			Set RSDIGG = Nothing
		    Set rst = oblog.Execute("select file_path,file_size,userid,fileid from oblog_upfile where logid=" & logid)
		    If Not rst.EOF Then
		        Set fs = CreateObject(oblog.CacheCompont(1))
		        Do While Not rst.EOF
			        fsize = rst(1)
			        uid = rst(2)
			        imgsrc = rst(0)
			        fid = rst(3)
			        If fs.FileExists(Server.MapPath(imgsrc)) Then
			            fs.DeleteFile (Server.MapPath(imgsrc))
			        End If

			        If InStr("jpg,bmp,gif,png,pcx", Right(imgsrc, 3)) > 0 Then '删除缩略图
			            imgsrc = Replace(imgsrc, Right(imgsrc, 3), "jpg")
			            imgsrc = Replace(imgsrc, Right(imgsrc, Len(imgsrc) - InStrRev(imgsrc, "/")), "pre" & Right(imgsrc, Len(imgsrc) - InStrRev(imgsrc, "/")))
			            If fs.FileExists(Server.MapPath(imgsrc)) Then
			                fs.DeleteFile Server.MapPath(imgsrc)
			            End If
			        End If
			        oblog.Execute ("delete from [oblog_upfile] where fileid=" & fid)
			        oblog.execute("update [oblog_user] set user_upfiles_size=user_upfiles_size-"&fsize&",user_upfiles_num=user_upfiles_num-1 where userid="&uid)
			        rst.Movenext
		        Loop
		        Set fs = Nothing
		        Set rst = Nothing
		    End If

		Next
	End Sub
End Class
%>