<%
Class class_sys
	Public Cache_Name, Cache_Name_Custom, Cache_data ,SqlQueryNum ,SqlQuery
	Public Reloadtime, setup, UserIp, ErrStr, AutoUpdate,CacheScores,CacheConfig,CacheCompont,CacheReport
	Public Userdir, User_CopyRight, ver, Is_password_cookies, defaultGroup
	Public l_uId, l_uName, l_uNickname,l_uPass, l_ulevel, l_uShowlogWord, l_uDir, l_isUbb, l_uDomain
	Public l_uFolder, l_uFrame,l_uGroupId,l_ucustomdomain,l_uUpUsed,l_uIco,l_uScores ,l_uNewBie,l_uAddtime
	Public l_uLastLogin,l_uLastComment,l_uLastMessage,l_uCommentCount,l_uMessageCount,l_uVisitCount,l_ulogcount
	Public l_Group,ShowBadWord,Time_Zone
	Public KeyWords1,KeyWords2,KeyWords3,KeyWords4
	Public NowUrl,Comeurl
	Public l_passport_userid ,l_is_log_default_hidden,l_blogpassword


	Private Sub Class_initialize()
		Reloadtime = 14400
		Cache_Name = blogdir & Cache_Name_user
		UserIp = GetIP
		Comeurl = LCase(Trim(Request.ServerVariables("HTTP_REFERER")))
		NowUrl = LCase(Trim(Request.ServerVariables("PATH_INFO")))
		ver = "4.60 Final"
		AutoUpdate = True			'更新整站首页开关
		Is_password_cookies = 0		'是否编码cookies,1为开启,0为关闭
		SqlQueryNum = 0
		Call ResetClassCache
	End Sub

	Private Sub class_terminate()
		On Error Resume Next
		If IsObject(conn) Then conn.Close: Set conn = Nothing
	End Sub

	Public Property Let name(ByVal vNewValue)
		Cache_Name_Custom = LCase(vNewValue)
	End Property

	Public Property Let Value(ByVal vNewValue)
		If Cache_Name_Custom <> "" Then
			ReDim Cache_data(2)
			Cache_data(0) = vNewValue
			Cache_data(1) = Now()
			Application.Lock
			Application(Cache_Name & "_" & Cache_Name_Custom) = Cache_data
			Application.unLock
		Else
			Err.Raise vbObjectError + 1, "CacheServer", " please change the CacheName."
		End If
	End Property

	Public Property Get Value()
		If Cache_Name_Custom <> "" Then
			Cache_data = Application(Cache_Name & "_" & Cache_Name_Custom)
			If IsArray(Cache_data) Then
				Value = Cache_data(0)
			Else
				Err.Raise vbObjectError + 1, "CacheServer", " The Cache_Data(" & Cache_Name_Custom & ") Is Empty."
			End If
		Else
			Err.Raise vbObjectError + 1, "CacheServer", " please change the CacheName."
		End If
	End Property

	Public Property Get SysDir()
		sysDir = Array ("admin","api","cam","data","editor","editor2","gg","images","inc","manager","oblogstyle","plus","skin","xmldata","xml-rpc")
	End Property
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
'oblog.ChkSpider(1) 1直接截断 0 输出检测结果
'------------------------------------------------
	Public Function  ChkSpider(tf)
		Dim ArrSpider,Agent_User,i
		'要屏蔽的搜索引擎的标志可以自己在下面加.
		ArrSpider=Array("google","baidu","yahoo","sina","sohu","jakarta","httpclient","soso","twiceler")
		Agent_User=LCase(Request.ServerVariables("HTTP_USER_AGENT"))
		ChkSpider = False
		For i = 0 To UBound(ArrSpider)
			If InStr(Agent_User,ArrSpider(i)) Then
				ChkSpider = True
				If tf=1 Then
					Response.Clear()
					Response.Status = 404
					Response.Write "<h1>404 Not Found</h1>"
					Response.End
				End If
						Exit Function
			End If
		Next

	End Function

	Public Function ObjIsEmpty()
		ObjIsEmpty = True
		Cache_data = Application(Cache_Name & "_" & Cache_Name_Custom)
		If Not IsArray(Cache_data) Then Exit Function
		If Not IsDate(Cache_data(1)) Then Exit Function
		If DateDiff("s", CDate(Cache_data(1)), Now()) < (60 * Reloadtime) Then ObjIsEmpty = False
	End Function

	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove (Cache_Name & "_" & MyCaheName)
		Application.unLock
	End Sub

	Public Sub ReloadSetup()
		Dim sql, rs, i
		sql = "select * from [oblog_setup] "
		Set rs = execute(sql)
		If rs.eof Then
			Response.Write "[oblog_setup]表信息不存在，无法正常运行程序！"
			Response.End
		End if
		name = "setup"
		Value = rs.GetRows(1)
		Set rs = Nothing
		Application.Lock
		Application(Cache_Name & "_index_update") = True
		Application(Cache_Name & "_list_update")  = True
		Application(Cache_Name & "_class_update") = False
		Application(Cache_Name & "_group_theme_main")=""
		Application(Cache_Name & "_Class_NeedUpdate")= True
		Application.unLock
	End Sub

	Public Sub ReLoadCache()
		Dim sql, rs, i,arr
		sql = "select * from oblog_config"
		Set rs = Execute(sql)
		If rs.eof Then
			Response.Write "[oblog_config]表信息不存在，无法正常运行程序！"
			Response.End
		End if
		Application.Lock
		rs.Filter="id=1"
		If Not rs.Eof Then
			arr=Split(rs(1),"$$")
		Else
			arr=""
		End If
		Application(Cache_Name & "_Config") = arr
		CacheConfig=Application(Cache_Name & "_Config")
		rs.Filter="id=2"
		If Not rs.Eof Then
			arr=Split(rs(1),"$$")
		Else
			arr=""
		End If
		Application(Cache_Name & "_Compont") = arr
		CacheCompont=Application(Cache_Name & "_Compont")
		rs.Filter="id=3"
		 If Not rs.Eof Then
			arr=Split(rs(1),"$$")
		Else
			arr=""
		End If
		Application(Cache_Name & "_Scores") = arr
		CacheScores=Application(Cache_Name & "_Scores")
		rs.Filter="id=4"
		 If Not rs.Eof Then
			arr=Split(rs(1),"$$")
		Else
			arr=""
		End If
		Application(Cache_Name & "_WhiteIp") = arr
		rs.Filter="id=5"
		 If Not rs.Eof Then
			arr=Split(rs(1),vbcrlf)
		Else
			arr=""
		End If
		Application(Cache_Name & "_BlackIp") = arr
		rs.Filter="id=6"
		If Not rs.Eof Then
			arr=Split(rs(1),vbcrlf)
		Else
			arr=""
		End If
		Application(Cache_Name & "_Keywords1") = arr
		KeyWords1= arr
		rs.Filter="id=7"
		If Not rs.Eof Then
			arr=Split(rs(1),vbcrlf)
		Else
			arr=""
		End If
		Application(Cache_Name & "_Keywords2") = arr
		KeyWords2= arr
		rs.Filter="id=8"
		If Not rs.Eof Then
			arr=Split(rs(1),vbcrlf)
		Else
			arr=""
		End If
		Application(Cache_Name & "_Keywords3") = arr
		KeyWords3= arr
		rs.Filter="id=9"
		 If Not rs.Eof Then
			arr=Split(rs(1),vbcrlf)
		Else
			arr=""
		End If
		Application(Cache_Name & "_Keywords4") = arr
		KeyWords4= arr
		rs.Filter="id=10"
		If Not rs.Eof Then
			arr=Split(rs(1),vbcrlf)
		Else
			arr=""
		End If
		Application(Cache_Name & "_Report") = arr
		CacheReport= arr
		Set rs=Execute("select top 1 Groupid From oblog_groups Order By g_level")
		Application(Cache_Name & "_defaultGroup") =rs(0)
		defaultGroup=Application(Cache_Name & "_defaultGroup")
		rs.Close
		Set rs=Nothing
		Application.unLock

	End Sub

	'读取用户目录及绑定的路径到缓存
	Public Sub ReloadUserdir()
		Dim sql, rs, s
		sql = "select userdir,dirdomain From oblog_userdir "
		Set rs = Execute(sql)
		While Not rs.EOF
			s = s & rs(0) & "!!??((" & rs(1) & "##))=="
			rs.movenext
		Wend
		Application.Lock
		Application(Cache_Name & "dirdomain") = s
		Application.unLock
		Set rs = Nothing
	End Sub

	Public Sub Start()
		CacheConfig=Application(Cache_Name & "_Config")
		CacheCompont=Application(Cache_Name & "_Compont")
		CacheScores=Application(Cache_Name & "_Scores")
		Keywords1=Application(Cache_Name & "_Keywords1")
		Keywords2=Application(Cache_Name & "_Keywords2")
		Keywords3=Application(Cache_Name & "_Keywords3")
		Keywords4=Application(Cache_Name & "_Keywords4")
		CacheReport=Application(Cache_Name & "_Report")
		defaultGroup=Application(Cache_Name & "_defaultGroup")
		name = "setup"
		If ObjIsEmpty() Then ReloadSetup()
		If Not IsArray(CacheConfig) Then ReLoadCache
		setup = Value
		'用户页面版权信息
		User_CopyRight = CacheConfig(7) & "</div>" & "<div id=""powered""><a href=""http://www.oblog.cn"" target=""_blank""><img src=""images\oblog_powered.gif"" border=""0"" alt=""Powered by "" /></a>"
		If DateDiff("s", Application(Cache_Name & "_index_updatetime"), Now()) > Int(CacheConfig(33)) And Application(Cache_Name & "_class_update") = True And AutoUpdate Then ReloadSetup()
		Time_Zone = Site_Time
	End Sub

	Public Sub Sys_Err(errmsg)
		ECHO_STR "出现一般系统错误","<b>产生错误的可能原因：</b><br>" & errmsg,1
	End Sub

	Public Function Site_bottom()
		Site_bottom = CacheConfig(10) & "<div id=""powered""><a href=""http://www.oblog.cn"" target=""_blank""><img src=""images\oblog_powered.gif"" border=""0"" alt=""Powered by "" /></a></div>"& vbCrLf
	End Function
	'获取服务器时区
	Function Site_Time()
	On Error Resume Next
		Dim intHours,ArrHours
		ArrHours=Split(oblog.CacheConfig(68),".")
		If UBound(ArrHours) = 0  Then
			intHours = oblog.CacheConfig(68)
		Else
			If Not IsNumeric(ArrHours(1)) Then
				intHours = ArrHours(0)
			Else
				intHours = oblog.CacheConfig(68)
			End if
		End If
		intHours =Int(FormatNumber(intHours,2))
		'防止未正确设置时区而抛出错误.
		If intHours="" Or IsNull(intHours) Then intHours = 8
		Site_Time = intHours
	End Function

	'------------------------------------------------
	'ServerDate(byval strDate)
	'服务器时差设置
	'回复/留言及发表日志
	'接收Trackback
	'------------------------------------------------
	Function ServerDate(byval strDate)
		Dim intHours
		strDate=Replace(strDate,"上午","AM")
		strDate=Replace(strDate,"下午","PM")
		strDate=Replace(strDate,"年","-")
		strDate=Replace(strDate,"月","-")
		strDate=Replace(strDate,"日","")
		strDate=Replace(strDate,".","-")
		If Not IsDate(strDate) Then
			ServerDate = Now()
			Exit Function
		End If
		'以北京时间为准
		intHours = Time_Zone - 8
		If Not IsNumeric(intHours) Then
			intHours = 0
			ServerDate = strDate
			Exit Function
		End If
		intHours =Int(intHours)
		If intHours > 24 Or intHours < -24 Then
			intHours = 0
			ServerDate=strDate
			Exit Function
		End If
		ServerDate = Dateadd("h",intHours,strDate)
	End Function

	Public Function Execute(SQL)
		If Not IsObject(CONN) Then link_database
		On Error Resume Next
'		Set Execute = conn.Execute(SQL)
		If InStr(LCase(SQL),"oblog_admin") Then


		End  If
		Dim Cmd
		Set Cmd = Server.CreateObject("ADODB.Command")
		Cmd.ActiveConnection = CONN
		Cmd.CommandText = SQL
		Set Execute = Cmd.Execute
		Set Cmd = Nothing
		If Err Then
			If Not Is_Debug Then
				Err.Clear
				ECHO_STR  "ExecuteSQL Err", "查询数据的时候发现错误，请检查您的查询代码是否正确。",0
			Else
				ECHO_STR "ExecuteSQL Err","查询数据的时候发现错误，请检查您的查询代码是否正确。<strong>ErrorSQL:</strong><br/>"&SQL&"<br /><br /><strong>Description:</strong>"&Err.Description ,0
			End If
			Set CONN = Nothing
			Response.End
		End if
		SqlQueryNum = SqlQueryNum + 1
		SqlQuery = SqlQuery & sql &"<br />"
	End Function


	Public Function chk_badword(Str)
		On Error Resume Next
		Dim badstr, i, n
		'先检查顶级过滤,如果存在则返回0.1
		'对于0.1情况需要特殊处理,0.1首先满足了>0的特点
		'但是对于日志发布时,如果是0.1,则列为可疑对象
		badstr = KeyWords1
		n = 0
		For i = 0 To UBound(badstr)
			If Trim(badstr(i)) <> "" Then
				If InStr(Str, Trim(badstr(i))) > 0 Then
					chk_badword = 0.1
					ShowBadWord = ShowBadWord & "," &Trim(badstr(i))
					Exit Function
				End If
			End If
		Next
		If ShowBadWord <> "" And Left(ShowBadWord,1)="," Then ShowBadWord =  Right (ShowBadWord,Len(ShowBadWord)-1)
		'检查审核过滤
		badstr = KeyWords2
		n = 0
		For i = 0 To UBound(badstr)
			If Trim(badstr(i)) <> "" Then
				If InStr(Str, Trim(badstr(i))) > 0 Then
					n = n + 1
				End If
			End If
		Next
		chk_badword = n
	End Function

	Public Function filt_badword(Str)
		On Error Resume Next
		Dim badstr, i
        badstr = KeyWords3
        For i = 0 To UBound(badstr)
            If Trim(badstr(i)) <> "" Then
                Str = Replace(Str, badstr(i), "***")
            End If
        Next
        filt_badword = Str
'		Dim objRegExp, strOutput,sKey
'		Set objRegExp = New Regexp
'		strOutput=Str
'		objRegExp.IgnoreCase = True
'		objRegExp.Global = True
'		badstr = KeyWords3
'		If UBound(badstr)=-1 Then
'			filt_badword=Str
'			Exit Function
'		End if
'		sKey=Join(badstr,"|")
'		objRegExp.Pattern = "(" & sKey & ")"
'		strOutput = objRegExp.replace(strOutput,"***")
'		filt_badword = strOutput
	End Function

	Public Function GetCode()
		Dim OBASN,CodeUrl ,Ist,isopen
		On Error Resume Next
		isopen=oblog.CacheConfig(85)
		If Err Then Err.clear:isopen=0
		Randomize
		OBASN=CStr(Int(900000*Rnd)+100000)
		CodeUrl = blogurl & IncCodePath & "?s="&OBASN
		ist= Not(Int(Right(OBASN,1)) = 0  Or Int(Right(OBASN,1)) = 6 ) And oblog.CacheConfig(85)=2
		If  isopen=0 Or right(split(LCase(Trim(Request.ServerVariables("SCRIPT_NAME"))),".asp")(0),5)="login" Or ist Then
		If Err Then Err.clear
			getcode = "<img id=""ob_codeimg"" src="""&CodeUrl&""" style=""cursor:hand;border:1px solid #ccc;vertical-align:top;"" onclick=""this.src='"&CodeUrl&"&t='+ Math.random();"" alt=""如果看不清数字或字母?请点一下换一个!"" title=""如果看不清数字或字母?请点一下换一个!"" /><input type=""hidden"" name=""ob_codename"" value="""&OBASN&""" /> " &vbcrlf


		ElseIf  isopen=1 Or isopen=2 Then
			getcode=getcode2(OBASN )

		End If
	End Function
'------(F)--------------生成并输出新的验证方式的验证
	Public Function GetCode2(OBASN)
		Dim CodeUrl
		Session("Ob_Ask_Shake_hands"&OBASN)=OBASN&"|" & "1"
		CodeUrl = blogurl & IncCodePath & "?n=2&s="&OBASN
		Rndcode(OBASN)
		GetCode2 = "<span id=""ob_codeimg"" onclick=""obaddjs('"&CodeUrl&"')"" alt=""如果看不懂问题或不知道怎么回答?请点一下换一个!"" title=""如果看不懂问题或不知道怎么回答?请点一下换一个!"" style=""cursor:hand;"">"&Session("OblogGetCode2_ask_"&OBASN)&"<br/>(请将答案填入验证码输入框.)</span><input type=""hidden"" name=""ob_codename"" value="""&OBASN&""" />"
	End Function

	Public Function Rndcode(OBASN)
	Dim sSql,rs
		sSql="select top 1 * From Oblog_Verifiydata "
		If CBool(Is_Sqldata) Then
			sSql= sSql & " Order By Newid()"
		Else
			Randomize
			sSql= sSql & " Order By Rnd(-(ID+"&Rnd()&"))"
		End If
		Set rs=oblog.Execute (sSql)
		If Not (rs.eof Or rs.bof) Then
		session("OblogGetCode2_ask_"&OBASN)=rs("ask")
		Session("GetCode"&OBASN)=rs("answer")
		Else
		session("OblogGetCode2_ask_"&OBASN)="随机问题数据库内没有随机问题数据!"
		Session("GetCode"&OBASN)=Empty
		End If
	End Function
'-------------------------------------------

	'检查验证码是否正确
	Public Function codepass()
		Dim CodeStr,codename,i,a,s
		CodeStr = Trim(Request("CodeStr"))
		codename = Trim(Request("ob_codename"))
		If LCase(CStr(Session("GetCode"&codename))) = LCase(CStr(CodeStr)) And CodeStr <> "" Then
			codepass = True
			Session("GetCode"&codename) = Empty
			Session("OblogGetCode2_ask_"&codename) = Empty
			Session("Ob_Ask_Shake_hands"&codename) = Empty
		ElseIf InStr(LCase(CStr(Session("GetCode"&codename))),"|")  And CodeStr <> ""  Then
			a=Split(LCase(CStr(Session("GetCode"&codename))),"|")
			For i=0 To UBound(a)
			If a(i) = LCase(CStr(CodeStr)) Then codepass = True
			Next
			Set a=Nothing
			Set i=Nothing
			Session("GetCode"&codename) = Empty
			Session("OblogGetCode2_ask_"&codename) = Empty
			Session("Ob_Ask_Shake_hands"&codename) = Empty
		ElseIf InStr(UCase(Session("GetCode")),";"&codename&":"&CodeStr&";") > 0 Then
			codepass = True
			s = UCase(Session("GetCode"))
			i = InStr(s,";"&codename&":")
			If i > 0 Then
				Session(GetCode) = Left(s,i) & Right(s,Len(s)-InStr(i+1,s,";"))
			End If
		Else
			codepass = False
			Session("GetCode"&codename) = Empty
			Session("OblogGetCode2_ask_"&codename) = Empty
		End If
	End Function

	Public Function type_domainroot(Str,sType)
		Dim domainroot, i
		If sType = 0 Then
			domainroot = Trim(cacheConfig(4))
		ElseIf sType = 1 Then
			domainroot = Trim(cacheConfig(75))
		End if
		If InStr(domainroot, "|") > 0 Then
			domainroot = Split(domainroot, "|")
			For i = 0 To UBound(domainroot)
				If Trim(domainroot(i)) <> "" Then
					If domainroot(i) = Str Then
					type_domainroot = type_domainroot & "<option value='" & Trim(domainroot(i)) & "' selected>" & "." & domainroot(i) & "</option>"
					Else
					type_domainroot = type_domainroot & "<option value='" & Trim(domainroot(i)) & "'>" & "." & domainroot(i) & "</option>"
					End If
				End If
			Next
		Else
			type_domainroot = "<option value='" & domainroot & "'>" & "." & domainroot & "</option>"
		End If
	End Function

	Public Function show_class(kind, CurrentID, kindType)
		If kind = "user" Then
			kind = 1
		Else
			kind = 2
		End if
		show_class=SelectedClassString(kind,kindType,CurrentID)
	End Function

	'取用户分类
	Public Function show_Postclass(CurrentID)
		show_Postclass=UserPostClass(2,0,CurrentID)
	End Function

	Public Sub AddErrStr(message)
		If errstr = "" Then
			errstr = message
		Else
			errstr = errstr & "_" & message
		End If
	End Sub

	Public Sub ShowErr()
		If errstr <> "" Then Response.Redirect blogurl & "err.asp?message=" & errstr
	End Sub

	Public Sub ShowUserErr()
		If errstr <> "" Then Response.Redirect blogurl & "user_prompt.asp?message=" & errstr
	End Sub

	Public Sub SaveCookie(username, password, CookieDate)
		Dim rs,userurl
		Set rs = oblog.Execute ("SELECT user_domain,user_domainroot,user_dir,user_folder FROM oblog_user WHERE username = '"&username&"' AND TruePassWord = '"&password&"' ")
		If rs.Eof Then Set rs = Nothing : Exit Sub
		If CacheConfig(4) <> "" And CacheConfig(5) = "1" Then
			'启用二级域名
			userurl = Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
		Else
			'未启用二级域名则从根目录开始访问，不包含域名
			userurl = blogdir & Trim(rs("user_dir")) & "/" & Trim(rs("user_folder")) & "/index." & f_ext
		End If
		select Case CookieDate
			Case 0
				'not save
			Case 1
				Response.Cookies(cookies_name).Expires = Date + 1
			Case 2
				Response.Cookies(cookies_name).Expires = Date + 31
			Case 3
				Response.Cookies(cookies_name).Expires = Date + 365
			Case Else
		End select
		If cookies_domain <> "" Then
			Response.Cookies(cookies_name).domain = cookies_domain
		End If
		Response.Cookies(cookies_name).Path   =   blogdir
		'不加密用户名,使登录的时候直接返回用户名.减少用户输入.
		Response.Cookies(cookies_name)("username") = username
		Response.Cookies(cookies_name)("password") = CodeCookie(password)
		Response.Cookies(cookies_name)("userurl") = CodeCookie(userurl)
	End Sub

	Public Sub ob_chklogin(username, password, CookieDate)
		Dim rs, sql ,TruePassWord,user_group,rsLogin,rsGroup
		Dim needUpdate,u_uid,u_gid
		needUpdate=False
		TruePassWord = RndPassword(16)
		If Not IsObject(conn) Then link_database
		Set rs = Server.CreateObject("adodb.recordset")
		sql = "select lockuser,userid,user_group,scores,TruePassWord,LastLoginIP,LastLoginTime,LoginTimes,user_domain,user_domainroot,user_dir,user_folder,"
		sql = sql & " user_upfiles_size"
		sql = sql & " FROM oblog_user "
		sql = sql & " WHERE username='" & username & "' AND password ='" & password & "' AND isdel=0 "
'		OB_Debug sql,1
		rs.open sql, conn, 1, 3
		If rs.EOF Then
			rs.Close: Set rs = Nothing
			adderrstr ("用户名或密码错误，请重新输入！")
			Exit Sub
		Else
			If rs("lockuser") = 1 Then
				rs.Close: Set rs = Nothing
				adderrstr ("对不起！你的ID已被锁定，不能登录！")
				Exit Sub
			Else
				'判断用户是否达到升级积分
				user_group = rs ("user_group")
				If IsNumeric(user_group) Then
					'获得组信息
					Set rsGroup = Execute ("select g_level FROM oblog_groups WHERE  groupid = "&user_group)
					If rsGroup.EOF Then
						ShowMsg "用户组信息不存在，请联系管理员",""
						Exit Sub
					End if
					'判断
					Set rsLogin=Execute("select top 1 groupid,g_points,g_autoupdate From oblog_groups Where g_level>" & CLng (rsGroup(0)) & " Order By g_level")
					If Not rsLogin.Eof Then
						If rsLogin(2)=1 Then
						'判断是否需要升级
							If rs("scores")>=Int(rsLogin(1)) Then
							needUpdate=True
							u_uid=rs("userid")
							u_gid=rsLogin(0)
							End If
						End If
					End If
				End If
				'基础防护，防止开启二级域名之后，之前的用户二级域名为空
				If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = "1" Then
					Dim user_domainroot,Arr_domainroot,TEMP_domainroot
					TEMP_domainroot=Trim(oblog.CacheConfig(4))
					If InStr(TEMP_domainroot,"|")>0 Then
						Arr_domainroot=Split(TEMP_domainroot,"|")
						user_domainroot=Arr_domainroot(0)
					Else
						user_domainroot=TEMP_domainroot
					End If
					rs("user_domain") = OB_IIF (rs("user_domain"),rs("userid"))
					rs("user_domainroot") = OB_IIF (rs("user_domainroot"),user_domainroot)
				End if
				'基础保护，登录时验证用户目录字段是否为空
				rs("user_dir") = OB_IIF (rs("user_dir"),setup(8,0))
				If CacheConfig (6) = "1" Then
					rs("user_folder") = OB_IIF (rs("user_folder"),username)
				Else
					rs("user_folder") = OB_IIF (rs("user_folder"),rs("userid"))
				End If
				rs("scores") = OB_IIF (rs("scores"),0)
				rs("user_upfiles_size") = OB_IIF (rs("user_upfiles_size"),0)
				rs("TruePassWord") = TruePassWord
				rs("LastLoginIP") = UserIp
				rs("LastLoginTime") = ServerDate(Now())
				rs("LoginTimes") = rs("LoginTimes") + 1
				rs.Update
				SaveCookie username, TruePassWord, CookieDate
				rsGroup.Close: Set rsGroup = Nothing
				rs.Close: Set rs = Nothing
			End If
		End If

		If needUpdate Then
								Execute ("update oblog_groups set g_members=g_members-1 WHERE groupid = " &user_group)
								Execute ("Update oblog_user Set user_group=" & u_gid & " Where userid=" & u_uid)
								Execute ("update oblog_groups set g_members=g_members+1 WHERE groupid = " &u_gid)
		End If
	End Sub
	'-------------------------------------
	Public Function CheckUserLogined_digg(puser,ppass)
		Dim rs
		If Not IsObject(conn) Then link_database
		Set rs = Server.CreateObject("adodb.recordset")
		rs.open "select top 1 userid,username from oblog_user where username='"&ProtectSQL(puser)&"' and truepassword='"&ProtectSQL(ppass)&"'", conn, 1, 1
		If Not (rs.eof Or rs.bof) Then
			CheckUserLogined_digg="1$$"&rs("userid")&"$$"&rs("username")
		Else
			CheckUserLogined_digg="0$$0$$0"
		End If
		rs.close
		Set rs=Nothing
	End Function
	'-------------------------------------
	Public Function CheckUserLogined()
		On Error Resume Next
		Dim Logined, rsLogin, sqlLogin, sSql, user_info ,tLogined ,i
		Logined = True
		'不加密用户名,使登录的时候直接返回用户名.减少用户输入.
		l_uName = filt_badstr(Request.Cookies(cookies_name)("username"))
		l_uPass = filt_badstr(DecodeCookie(Request.Cookies(cookies_name)("password")))
		If l_uName = "" Then
			Logined = False
		End If
		If l_uPass = "" Then
			Logined = False
		End If
		sSql = "userid,user_level,user_showlogword_num,user_upfiles_max,user_upfiles_size,user_dir,isubbedit,user_domain,"
		sSql = sSql &"user_domainroot,lockuser,user_folder,adddate,user_info,user_Icon1,user_Icon2,user_group,lastcomment,"
		sSql = sSql &"lastmessage,scores,Nickname,comment_count,message_count,newbie,lastlogintime,log_count,user_siterefu_num,passport_userid,is_log_default_hidden,blog_password" & str_domain
		If Logined = True Then
			If Session ("CheckUserLogined_"&l_uName) = "" Then
				'除了str_domain，0-28列
				sqlLogin = "select " & sSql & " from oblog_user where Username='" & l_uName & "' and TruePassword='" & l_uPass & "' "
				Set rsLogin = Execute(sqlLogin)

				If rsLogin.EOF Then
					CheckUserLogined = false
					Exit Function
				Else
					If rsLogin(9) = 1 Or IsNull( rsLogin(9)) Then
						Set rsLogin = Nothing
						adderrstr ("当前用户已被系统锁定，无法进行操作，请联系管理员！")
						showerr
						Exit Function
					End If
					For i = 0 To 28
						tLogined = tLogined & "$$$" & rsLogin(i)
					Next
					tLogined = Right (tLogined,Len(tLogined)-3)
					If str_domain <> "" Then tLogined = tLogined & "$$$" &rsLogin("custom_domain")
					Session ("CheckUserLogined_"&l_uName) = tLogined
				End If
			End If
			tLogined = Session ("CheckUserLogined_"&l_uName)
			tLogined = Split (tLogined,"$$$")
'			Response.Write tLogined(18)
'			Response.Write tLogined(19)
			'Response.Write "|"&tLogined(28)&"|"
			'Response.Write UBound(tLogined)
			If UBound(tLogined) > 29 Or UBound(tLogined) = 0 Or UBound(tLogined) = -1 Then
				Session ("CheckUserLogined_"&l_uName) = ""
				Response.Redirect (blogurl & "login.asp")
				Exit Function
			End if
			l_uId = Int(tLogined(0))
			l_ulevel = Int(tLogined(1))
			l_uShowlogWord = Int(tLogined(2))
			l_uDir = tLogined(5)
			l_isUbb = Int(OB_IIF(tLogined(6),2))
			l_uDomain = tLogined(7) & "." & tLogined(8)
			l_uFolder = tLogined(10)
			l_uGroupId=Int(OB_IIF(tLogined(15),1))
			l_uUpUsed=Int(tLogined(4))
			l_uLastComment=tLogined(16)
			l_uLastMessage=tLogined(17)
			l_uScores=Int(OB_IIF(tLogined(18),100))
			l_uNickname=tLogined(19)
			l_uCommentCount=Int(OB_IIF(tLogined(20),0))
			l_uMessageCount=Int(OB_IIF(tLogined(21),0))
			l_uNickname = ob_IIF(l_uNickname,l_uName)
			If InStr(tLogined(11), "$") Then
				user_info = Split(tLogined(11), "$")
				l_uFrame = user_info(1)
			Else
				l_uFrame = 1
			End If

			If tLogined(28)<>""  Then
			l_blogpassword=1
			Else
			l_blogpassword=0
			End If

			If true_domain = 1 Then
				'判断用户绑定的顶级域名
				l_ucustomdomain = tLogined(29)
				If l_ucustomdomain <> "" Then
					l_uDomain = l_ucustomdomain
				End If
			End If

			l_is_log_default_hidden=OB_IIF(tLogined(27),0)
			If Err Then Err.clear:l_is_log_default_hidden=0
			l_passport_userid=OB_IIF(tLogined(26),0)
			l_uNewBie=Int(OB_IIF(tLogined(22),0))
			l_uIco=ProIco(tLogined(13), 1)
			l_uLastLogin=tLogined(23)
			l_ulogcount=Int(tLogined(24))
			l_uvisitcount=Int(tLogined(25))
			l_uAddtime=tLogined(11)
			Call GetGroupInfo
			Set rsLogin = Nothing
		End If

		If l_isUbb > 0 Then C_Editor_Type = l_isUbb
		Select Case C_Editor_Type
			Case 1
				C_Editor=blogdir&"editor"
				C_Editor_LoadIcon="yes"
			Case 2
				C_Editor=blogdir&"editor2"
				C_Editor_LoadIcon="none"
		End Select
		C_Editor_UBB=blogurl&"editor"
		If Err Then
			Err.Clear
			Session ("CheckUserLogined_"&l_uName) = ""
			Logined = False
			Response.Redirect (blogurl & "index.asp")
		End If
		CheckUserLogined = Logined
	End Function
	'组信息
	Public Sub GetGroupInfo()
		Dim rst
		Set rst=Execute("select * From oblog_groups Where groupid=" & CLng (l_uGroupId) )
		If Not rst.Eof Then
			l_Group=rst.GetRows(1)
		Else
			ShowMsg "用户组信息不存在，请联系管理员","index.asp"
		End If
		Set rst=Nothing
	End Sub

	Public Sub CreateUserDir(ustr, action)
	On Error Resume Next
		Dim fso, sql, rs, udir, uid, upath, loginstr, searchstr, bname, ufolder, utruepath,uname
		sql = "select userid,user_dir,blogname,user_folder,username,user_domain,user_domainroot" & str_domain & " from oblog_user where "
		If action = 0 Then sql = sql & "userid=" & CLng (ustr) Else sql = sql & "username='" & filt_badstr(ustr) & "'"
		Set rs = Execute(sql)
		If Not rs.EOF Then
			udir = rs(1)
			uid = rs(0)
			bname = rs(2)
			ufolder = rs(3)
			uname = rs(4)
			'基础保护，防止无法生成用户页面.
			If udir = "" Or IsNull(udir) Then
				udir = ob_iif(setup(8,0),"u")
				Execute ("UPDATE oblog_user SET user_dir = '"&udir&"' WHERE userid = " &uid)
			End If
			If ufolder = "" Or IsNull(ufolder) Then
				If CacheConfig (6) = "1" Then
					ufolder = uid
				Else
					ufolder = uname
				End If
				'过滤含有.的用户目录,防止由用户通过其它方式该用户目录为asp.asp
				If InStr(LCase(ufolder),".") Then ufolder = uid
				Execute ("UPDATE oblog_user SET user_folder = '"&ufolder&"' WHERE userid = " &uid)
			End If
			If true_domain = 1 Then
				If rs("custom_domain") <> "" And Not IsNull(rs("custom_domain")) Then
					utruepath = "http://" & rs("custom_domain") & "/"
				Else
					utruepath = "http://" & rs("user_domain") & "." & rs("user_domainroot") & "/"
				End If
			Else
				utruepath = blogdir & udir & "/" & ufolder & "/"
			End If
			If bname = "" Or IsNull(bname) Then bname = " "
			searchstr = "<form name=""search"" method=""post"" action=""" & blogurl & "list.asp?userid=" & uid & """ target=""_blank"">" & vbcrlf
			searchstr = searchstr & "	<select name=""selecttype"" id=""selecttype"">" & vbcrlf
			searchstr = searchstr & "		<option value=""topic"" selected>日志标题</option>" & vbcrlf
			searchstr = searchstr & "		<option value=""logtext"">日志内容</option>" & vbcrlf
			searchstr = searchstr & "	</select>" & vbcrlf
			searchstr = searchstr & "	<br />" & vbcrlf
			searchstr = searchstr & "	<input name=""keyword"" type=""text"" id=""keyword"" size=""16"" maxlength=""40"">" & vbcrlf
			searchstr = searchstr & "	<input type=""submit"" name=""Submit"" value=""搜索"">" & vbcrlf
			searchstr = searchstr & "</form>" & vbcrlf

			'upath = Server.MapPath(udir)
			upath = Server.MapPath(blogdir & udir)
			Set fso = Server.CreateObject(CacheCompont(1))
			If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
			upath = Server.MapPath(blogdir & udir & "/" & ufolder)
			If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
			Call BuildFile(upath & "/index." & f_ext, "暂无日志,请发表日志或者更新首页!" )
			Call BuildFile(upath & "/message." & f_ext, "暂无留言,请更新发布留言板!" )
			upath = Server.MapPath(blogdir & udir & "/" & ufolder & "/calendar")
			If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
			If f_ext = "htm" Or f_ext = "html" Then
				Call BuildFile(upath & "/0.htm", htm2js_div(" ", "calendar") )
			Else
				Call BuildFile(upath & "/0.htm", " " )
			End If
			upath = Server.MapPath(blogdir & udir & "/" & ufolder)
			If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
			Dim xmlDoc,userpath
			Set xmlDoc = New Cls_XmlDoc
			userpath = blogdir & udir & "/" & ufolder&"/user.config"
			If xmlDoc.LoadXml (blogdir&"XmlData/user.config") Then
				xmlDoc.SaveAs userpath
			Else
				Response.Write (blogdir&"XmlData/user.config 不存在，无法继续操作！")
				Set XmlDoc = Nothing
				Response.End
			End If
			If xmlDoc.LoadXml (userpath) Then
				xmlDoc.UpdateNodeText "blogname",oblog.htm2js_div(filt_html(bname),"blogname"),True
				xmlDoc.UpdateNodeText "placard",oblog.htm2js_div(" ","placard"),True
				xmlDoc.UpdateNodeText "subject",oblog.htm2js_div(" ","subject"),True
				xmlDoc.UpdateNodeText "newblog",oblog.htm2js_div(" ","newblog"),True
				xmlDoc.UpdateNodeText "comment",oblog.htm2js_div(" ","comment"),True
				xmlDoc.UpdateNodeText "links",oblog.htm2js_div(" ","links"),True
				xmlDoc.UpdateNodeText "info",oblog.htm2js_div(" ","info"),True
				xmlDoc.UpdateNodeText "search",oblog.htm2js_div(searchstr,"search"),True
				xmlDoc.UpdateNodeText "mygroups",oblog.htm2js_div(" ","mygroups"),True
				xmlDoc.UpdateNodeText "myfriend",oblog.htm2js_div(" ","myfriend"),True
				xmlDoc.UpdateNodeText "newmessage",oblog.htm2js_div("<a href=""" & utruepath & "message." & f_ext & "#cmt""><strong>签写留言</strong></a> ","newmessage"),True
				xmlDoc.Save
				Set xmlDoc = Nothing
			Else
				Response.Write xmlDoc.ErrInfo
				Set xmlDoc = Nothing
				Response.End
			End if
			If CacheConfig(57) = "1" Then
				upath = Server.MapPath(blogdir & udir & "/" & ufolder & "/archives")
				If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
			End If
			Set fso = Nothing
			Set rs = Nothing
		Else
			Set rs = Nothing
			Response.Write ("没找到该用户，无法建立目录。")
			Exit Sub
		End If
	End Sub

	Public Sub ShowMsg(Str, url)
		url = LCase(Trim(url))
		If url = "" Then
			'如果返回URL为空
			'如果可以获取来路则直接返回来路，否则返回上一页
			If Comeurl = "" Then
				Response.Write "<script language=Javascript>alert(""" & Str & """);history.go(-1)</script>"
			Else
				Response.Write "<script language=Javascript>alert(""" & Str & """);window.location='" & Comeurl & "'</script>"
			End if
		Else
			'操作完成后关闭当前窗口
			If url = "close" Then
				Response.Write "<script language=Javascript>alert(""" & Str & """);self.close();</script>"
			ElseIf url="back" Then
				Response.Write "<script language=Javascript>alert(""" & Str & """);history.back()</script>"
			Else
			'操作完成后转向目标URL
				Response.Write "<script language=Javascript>alert(""" & Str & """);window.location='" & url & "'</script>"
			End if
		End If
		Set oblog = Nothing
		Response.End
	End Sub

	Public Function type_city(province, city)
		Dim tmpstr
		tmpstr = "	<select onchange=setcity(); name=""province"">" & vbcrlf
		tmpstr = tmpstr & "		<option value="""">选择省份</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""安徽"">安徽</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""北京"">北京</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""重庆"">重庆</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""福建"">福建</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""甘肃"">甘肃</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""广东"">广东</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""广西"">广西</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""贵州"">贵州</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""海南"">海南</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""河北"">河北</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""黑龙江"">黑龙江</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""河南"">河南</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""香港"">香港</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""湖北"">湖北</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""湖南"">湖南</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""江苏"">江苏</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""江西"">江西</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""吉林"">吉林</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""辽宁"">辽宁</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""澳门"">澳门</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""内蒙古"">内蒙古</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""宁夏"">宁夏</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""青海"">青海</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""山东"">山东</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""上海"">上海</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""山西"">山西</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""陕西"">陕西</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""四川"">四川</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""台湾"">台湾</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""天津"">天津</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""新疆"">新疆</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""西藏"">西藏</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""云南"">云南</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""浙江"">浙江</option>" & vbcrlf
		tmpstr = tmpstr & "		<option value=""海外"">海外</option>" & vbcrlf
		tmpstr = tmpstr & "	</select>" & vbcrlf
		tmpstr = tmpstr & "	<select name=""city"">" & vbcrlf
		tmpstr = tmpstr & "	</select>" & vbcrlf
		tmpstr = tmpstr & "<script src=""inc/getcity.js""></script>" & vbcrlf
		tmpstr = tmpstr & "<script>initprovcity('" & province & "','" & city & "');</script>" & vbcrlf
		type_city = tmpstr
	End Function
	Public Sub type_job(job)
		Dim tmpstr
		tmpstr = "<select name=""job"" id=""job"">" & vbcrlf
		tmpstr = tmpstr & "	<option value="""">----请选择职业----</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""财会/金融""> 财会/金融</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""工程师"">工程师</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""顾问"">顾问</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""计算机相关行业"">计算机相关行业</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""计算机相关行业（其他）"">计算机相关行业（其他）</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""家庭主妇"">家庭主妇</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""教育/培训"">教育/培训</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""客户服务/支持"">客户服务/支持</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""零售商/手工工人"">零售商/手工工人</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""退休"">退休</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""无职业"">无职业</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""销售/市场/广告"">销售/市场/广告</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""学生"">学生</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""研究和开发"">研究和开发</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""一般管理"">一般管理</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""政府/军队"">政府/军队</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""执行官/高级管理"">执行官/高级管理</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""制造/生产/操作"">制造/生产/操作</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""专业人员（医药、法律等）"">专业人员（医药、法律等）</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""自雇/业主"">自雇/业主</option>" & vbcrlf
		tmpstr = tmpstr & "	<option value=""其他"">其他</option>" & vbcrlf
		tmpstr = tmpstr & "</select>" & vbcrlf
		Response.Write (tmpstr)
	%>
	<script language=javascript>
	var jobObject = document.oblogform["job"];
	for(var i = 0; i < jobObject.options.length; i++) {
		if (jobObject.options[i].value=="<%=Trim(job)%>")
		{
			jobObject.selectedIndex = i;
		}
	}
	</script>
	<%
	End Sub

	Public Sub type_dateselect(addtime, n)
		Dim y, m, d, ttime
		If addtime = "" Then ttime = ServerDate(Now()) Else ttime = addtime
		Response.Write("<select name=""selecty"&n&""">")&vbcrlf
		For y = Year(Now())-10 To Year(Now())+10
			If Year(ttime) = y Then
				Response.Write "<option value="""&y&""" selected>"&y&"年</option>"&vbcrlf
			Else
				Response.Write "<option value="""&y&""">"&y&"年</option>"&vbcrlf
			End If
		Next
		Response.Write "</select>"&vbcrlf
		Response.Write "<select name=""selectm"&n&""">"&vbcrlf

		For m = 1 To 12
			If Month(ttime) = m Then
				Response.Write "<option value="""&m&""" selected>"&m&"月</option>"&vbcrlf
			Else
				Response.Write "<option value="""&m&""">"&m&"月</option>"&vbcrlf
			End If
		Next
		Response.Write("</select>")&vbcrlf
		Response.Write("<select name=""selectd"&n&""">")&vbcrlf

		For d = 1 To 31
			If Day(ttime) = d Then
				Response.Write "<option value="""&d&""" selected>"&d&"日</option>"&vbcrlf
			Else
				Response.Write "<option value="""&d&""">"&d&"日</option>"&vbcrlf
			End If
		Next
		Response.Write ("</select>") & vbCrLf
	End Sub

	Public Sub chk_commenttime()
		Dim lasttime
		if CacheConfig(27) = "0" Then
			If DateDiff("s", l_uLastComment, l_uLastMessage) > 0 Then
				lasttime = l_uLastMessage
			Else
				lasttime = l_uLastComment
			End If
		Else
			lasttime = Request.Cookies(cookies_name)("LastComment")
		End If
		If IsDate(lasttime) Then
			If DateDiff("s", lasttime, ServerDate(Now())) < Int(cacheConfig(32)) Then
				Response.Write ("<script language=javascript>alert('" & cacheConfig(32) & "秒后才能回复或评论。');window.history.back(-1);</script>")
				Response.End
			End If
		End If
	End Sub

	Public Function filtpath(Str)
		Dim s1
		If oblog.CacheConfig(55) = 1 Then
			Dim nurl
			nurl = Trim("http://" & Request.ServerVariables("HTTP_HOST"))
			nurl = nurl & Request.ServerVariables("PATH_INFO")
			nurl = Left(nurl, InStrRev(nurl, "/"))
			s1 = Replace(Str, nurl, "")
		Else
			s1 = Str
		End If
		filtpath=Replace(s1,"over--flow","overflow")
	End Function


	Public Function showpage(bTotal, bAllPages, sUnit)
		Dim n, i, sTmp, strUrl
		If G_P_PerMax=0 Then G_P_PerMax=1
		If G_P_AllRecords Mod G_P_PerMax = 0 Then
			n = G_P_AllRecords \ G_P_PerMax
		Else
			n = G_P_AllRecords \ G_P_PerMax + 1
		End If
		sTmp = vbcrlf & "<div id=""showpage"">" & vbcrlf
		If bTotal = True Then
			sTmp = sTmp & "共" & G_P_AllRecords & sUnit & "&nbsp;&nbsp;"
		End If
		strUrl = JoinChar(G_P_FileName)
		If G_P_This < 2 Then
				sTmp = sTmp & "首页 上一页&nbsp;"
		Else
				sTmp = sTmp & "<a href=""" & strUrl & "page=1"">首页</a>&nbsp;"
				sTmp = sTmp & "<a href=""" & strUrl & "page=" & (G_P_This - 1) & """>上一页</a>&nbsp;"
		End If

		If n - G_P_This < 1 Then
				sTmp = sTmp & "下一页 尾页"
		Else
				sTmp = sTmp & "<a href=""" & strUrl & "page=" & (G_P_This + 1) & """>下一页</a>&nbsp;"
				sTmp = sTmp & "<a href=""" & strUrl & "page=" & n & """>尾页</a>"
		End If
		sTmp = sTmp & "&nbsp;页次：" & G_P_This & "/" & n & "页 "
		sTmp = sTmp & "&nbsp;" & G_P_PerMax & "" & sUnit & "/页"
		If bAllPages = True Then
			sTmp = sTmp & "&nbsp;转到：<select name=""page"" size=""1"" onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"
			For i = 1 To n
				sTmp = sTmp & "<option value=""" & i & """"
				If CInt(G_P_This) = CInt(i) Then sTmp = sTmp & " selected "
				sTmp = sTmp & ">" & i & "</option>"
			Next
			sTmp = sTmp & "</select>"
		End If
		sTmp = sTmp & "</div>" & vbcrlf
		showpage = sTmp
	End Function

	Function MakePageBar(rs,sUnit)
		if Request("page")<>"" then
			G_P_This=cint(Request("page"))
		else
			G_P_This=1
		end if
		If rs.EOF Then
			G_P_Guide = G_P_Guide & " (共有0"&sUnit&")"
			Response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & G_P_Guide
		Else
			G_P_AllRecords = rs.recordcount
			G_P_Guide = G_P_Guide & " (共有" & G_P_AllRecords & sUnit & ")"
			If G_P_This < 1 Then
				G_currentPage = 1
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
				Response.write oblog.showpage(True, True, sUnit)
			Else
				If (G_P_This - 1) * G_P_PerMax < G_P_AllRecords Then
					rs.Move (G_P_This - 1) * G_P_PerMax
					Dim bookmark
					bookmark = rs.bookmark
					showContent
					Response.write oblog.showpage(True, True, sUnit)
				Else
					G_currentPage = 1
					showContent
					Response.write oblog.showpage(True, True, sUnit)
				End If
			End If
		End If
	End Function
	Public Function JoinChar(strUrl)
		If strUrl = "" Then
			JoinChar = ""
			Exit Function
		End If
		If InStr(strUrl, "?") < Len(strUrl) Then
			If InStr(strUrl, "?") > 1 Then
				If InStr(strUrl, "&") < Len(strUrl) Then
					JoinChar = strUrl & "&"
				Else
					JoinChar = strUrl
				End If
			Else
				JoinChar = strUrl & "?"
			End If
		Else
			JoinChar = strUrl
		End If
	End Function

	Public Function htm2js(Str,IsWrite)
		If Str = "" Or IsNull(Str) Then Str = " "
		Str = Replace(Str, "\", "\\")
		Str = Replace(Str, "'", "\'")
'		Str = Replace(Str, vbCrLf, "\n")
		Str = Replace(Str, Chr(13), "")
		Str = Replace(Str, Chr(10), "\n")
		If IsWrite Then
			htm2js = "document.write('" & Str & "');"
		Else
			htm2js = Str
		End If
	End Function

	'将htm代码插入div,不支持脚本插入
	Public Function htm2js_div(Str, divid)
		divid = Trim(divid)
		If Str = "" Or IsNull(Str) Then Str = " "
		Str = Replace(Str, "\", "\\")
		Str = Replace(Str, "'", "\'")
'		Str = Replace(Str, vbCrLf, "\n")
		Str = Replace(Str, Chr(13), "")
		Str = Replace(Str, Chr(10), "")
		htm2js_div = "if (chkdiv('" & divid & "')) {"
		htm2js_div = htm2js_div & "document.getElementById('" & divid & "')" & ".innerHTML='" & Str & "';}"
		If divid = "subject" Then htm2js_div = htm2js_div & vbCrLf & "if (chkdiv('subject_l')) {document.getElementById('subject_l').innerHTML='" & Str & "';}"
	End Function

	'将htm代码插入div,支持脚本插入
	'效率低下，除非必须，否则不建议使用
	Public Function htm2js_Script(Str, divid)
		divid = Trim(divid)
		If Str = "" Or IsNull(Str) Then Str = " "
		Str = Replace(Str, "\", "\\")
		Str = Replace(Str, "'", "\'")
'		Str = Replace(Str, vbCrLf, "\n")
		Str = Replace(Str, Chr(13), "")
		Str = Replace(Str, Chr(10), "\n")
		htm2js_Script = "if (chkdiv('" & divid & "')) {"
		htm2js_Script = htm2js_Script & "set_innerHTML('" & divid & "','" & Str & "');}"
	End Function

	Public Function readfile(mPath, fName)
		On Error Resume Next
		Dim fs2, f2, fpath
		fpath = Server.MapPath(mPath) & "\"
		fpath = fpath & fName
		If CacheConfig(24) = "1" Then
			Dim oStream
			Set oStream = Server.CreateObject(CacheCompont(2))
			With oStream
				.Type = 2
				.Mode = 3
				.open
				'.Charset = "utf-8"
				.Charset = "gb2312"
				.Position = oStream.size
				.open
				.loadfromfile fpath
			End With
			readfile = oStream.readtext
			oStream.Close
			Set oStream = Nothing
		Else
 			Set fs2 = Server.CreateObject(CacheCompont(1))
			Set f2 = fs2.OpenTextFile(fpath, 1, True)
			readfile = f2.ReadAll
			Set fs2 = Nothing
			Set f2 = Nothing
		End If
	End Function

	Public Function showsize(ByVal size)
		On Error Resume Next
		If size = "" Or IsNull(size) Then
			showsize = "0Byte"
			Exit Function
		End If
		showsize = size & "Byte"
		If size < 0 Then
			showsize = "0KB"
			Exit Function
		End If
		If size > 1024 Then
		   size = (size / 1024)
		   showsize = FormatNumber(size, 2) & "KB"
		End If
		If size > 1024 Then
		   size = (size / 1024)
		   showsize = FormatNumber(size, 2) & "MB"
		End If
		If size > 1024 Then
		   size = (size / 1024)
		   showsize = FormatNumber(size, 2) & "GB"
		End If
		If size > 1024 Then
		   size = (size / 1024)
		   showsize = FormatNumber(size, 2) & "TB"
		End If
		If size > 1024 Then
		   size = (size / 1024)
		   showsize = FormatNumber(size, 2) & "PB"
		End If
		If size > 1024 Then
		   size = (size / 1024)
		   showsize = FormatNumber(size, 2) & "EB"
		End If
	End Function

	Public Function ChkPost()
		Dim server_v1, server_v2
		ChkPost = False
		If true_domain = 1 Then
			ChkPost = True
			Exit Function
		End If
		server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
		server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
		If server_v1 = GetUrl Then
'			Exit Function
		End If
		If Mid(server_v1, 8, Len(server_v2)) = server_v2 Then ChkPost = True
	End Function

	Public Function filt_badstr(sSql)
		 If IsNull(sSql) Then Exit Function
		 sSql = Trim(sSql)
		 If sSql = "" Then Exit Function
		 sSql = Replace(sSql, Chr(0), "")
		 sSql = Replace(sSql, "'", "''")
		 'sSql=Replace(sSql,"%","％")
		 'sSql=Replace(sSql,"-","－")
		 filt_badstr = sSql
	End Function

	Public Function filt_astr(Str, n)
		If IsNull(Str) Then
			filt_astr = ""
			Exit Function
		End If
		filt_astr = filt_badword(Str)
		filt_astr = InterceptStr(filt_astr, n)
	End Function

	Public Function filt_html(Str)
		On Error Resume Next
		If Str = "" Then
			filt_html = ""
		Else
			If IsNull(Str) Then
				filt_html = Str
				Exit Function
			End if
			Str = Replace(Str, ">", "&gt;")
			Str = Replace(Str, "<", "&lt;")
			Str = Replace(Str, Chr(32), "&nbsp;")
			Str = Replace(Str, Chr(9), "&nbsp;")
			Str = Replace(Str, Chr(34), "&quot;")
			Str = Replace(Str, Chr(39), "&#39;")
			Str = Replace(Str, Chr(13), "")
			Str = Replace(Str, Chr(10) & Chr(10), "&nbsp; ")
			Str = Replace(Str, Chr(10), "&nbsp; ")
			filt_html = Str
		End If
	End Function

	Public Function filt_html_b(fString)
		On Error Resume Next
		If Not IsNull(fString) And fString<>"" Then
			fString = Replace(fString, ">", "&gt;")
			fString = Replace(fString, "<", "&lt;")
			fString = Replace(fString, Chr(32), " ")
			fString = Replace(fString, Chr(9), " ")
			fString = Replace(fString, Chr(34), "&quot;")
			'fString = Replace(fString, CHR(39), "&#39;")
			fString = Replace(fString, Chr(13), "")
			fString = Replace(fString, Chr(10) & Chr(10), "</p><p> ")
			fString = Replace(fString, Chr(10), "<br> ")
			filt_html_b = fString
		Else
			filt_html_b=""
		End If
	End Function

	Public Function strLength(Str)
		On Error Resume Next
		Dim WINNT_CHINESE
		WINNT_CHINESE = (Len("中国") = 2)
		If WINNT_CHINESE Then
			Dim l, t, c
			Dim i
			l = Len(Str)
			t = l
			For i = 1 To l
				c = Asc(Mid(Str, i, 1))
				If c < 0 Then c = c + 65536
				If c > 255 Then
					t = t + 1
				End If
			Next
			strLength = t
		Else
			strLength = Len(Str)
		End If
		If Err.Number <> 0 Then Err.Clear
	End Function

	Public Function InterceptStr(txt, length)
		On Error Resume Next
		Dim WINNT_CHINESE
		WINNT_CHINESE = (Len("中国") = 2)
		If WINNT_CHINESE Then InterceptStr = Left (txt,length):Exit Function
		Dim x, y, ii
		txt = Trim(txt)
		x = Len(txt)
		y = 0
		If x >= 1 Then
			For ii = 1 To x
				If Asc(Mid(txt, ii, 1)) < 0 Or Asc(Mid(txt, ii, 1)) > 255 Then '如果是汉字
					y = y + 2
				Else
					y = y + 1
				End If
				If y >= length Then
					txt = Left(Trim(txt), ii) '字符串限长
					Exit For
				End If
			Next
			InterceptStr = txt
		Else
			InterceptStr = ""
		End If
	End Function

	Public Function GetUrl()
		On Error Resume Next
		Dim sTmp
		If LCase(Request.ServerVariables("HTTPS")) = "off" Then
			sTmp = "http://"
		Else
			sTmp = "https://"
		End If
		sTmp = sTmp & Request.ServerVariables("SERVER_NAME")
		If Request.ServerVariables("SERVER_PORT") <> 80 Then sTmp = sTmp & ":" & Request.ServerVariables("SERVER_PORT")
		sTmp = sTmp & Request.ServerVariables("PATH_INFO")
		If Trim(Request.QueryString) <> "" Then sTmp = sTmp & "?" & Trim(Request.QueryString)
		GetUrl = sTmp
	End Function

	Public Function trueurl(strContent)
		On Error Resume Next
		Dim tempReg, url
		url = Trim("http://" & Request.ServerVariables("HTTP_HOST"))
		url = LCase(url & Request.ServerVariables("PATH_INFO"))
		url = Left(url, InStrRev(url, "/"))
		Set tempReg = New RegExp
		tempReg.IgnoreCase = True
		tempReg.Global = True
		tempReg.Pattern = "(^.*\/).*$" '含文件名的标准路径
		url = tempReg.replace(url, "$1")
		tempReg.Pattern = "((?:src|href).*?=[\'\u0022](?!ftp|http|https|mailto))"
		trueurl = tempReg.replace(strContent, "$1" + url)
		Set tempReg = Nothing
	End Function

	Public Function IsValidEmail(email)
		Dim names, name, i, c
		IsValidEmail = True
		names = Split(email, "@")
		If UBound(names) <> 1 Then
		   IsValidEmail = False
		   Exit Function
		End If
		For Each name In names
		   If Len(name) <= 0 Then
			 IsValidEmail = False
			 Exit Function
		   End If
		   For i = 1 To Len(name)
			 c = LCase(Mid(name, i, 1))
			 If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
			   IsValidEmail = False
			   Exit Function
			 End If
		   Next
		   If Left(name, 1) = "." Or Right(name, 1) = "." Then
			  IsValidEmail = False
			  Exit Function
		   End If
		Next
		If InStr(names(1), ".") <= 0 Then
		   IsValidEmail = False
		   Exit Function
		End If
		i = Len(names(1)) - InStrRev(names(1), ".")
		If i <> 2 And i <> 3 Then
		   IsValidEmail = False
		   Exit Function
		End If
		If InStr(email, "..") > 0 Then
		   IsValidEmail = False
		End If
	End Function
	'只允许数字(48~57)+大(65~90)小(97~122)写字母和下划线
	Public Function chkDomain(domain)
		Dim name, i, c
		name = domain
		chkdomain = True
		If Len(name) <= 0 Then
			chkdomain = False
			Exit Function
		End If
		For i = 1 To Len(name)
			c = LCase(Mid(name, i, 1))
			If InStr("abcdefghijklmnopqrstuvwxyz_", c) <= 0 And Not IsNumeric(c) Then
				chkdomain = False
				Exit Function
			End If
		Next
	End Function

	Public Function CodeCookie(Str)
		If Is_password_cookies = 1 Then
			Dim i
			Dim StrRtn
			For i = Len(Str) To 1 Step -1
				StrRtn = StrRtn & AscW(Mid(Str, i, 1))
				If (i <> 1) Then StrRtn = StrRtn & "a"
			Next
			CodeCookie = StrRtn
		Else
			CodeCookie = Str
		End If
	End Function

	Public Function DecodeCookie(Str)
		If Is_password_cookies = 1 Then
			Dim i
			Dim StrArr, StrRtn
			StrArr = Split(Str, "a")
			For i = 0 To UBound(StrArr)
				If IsNumeric(StrArr(i)) = True Then
					StrRtn = ChrW(StrArr(i)) & StrRtn
				Else
					StrRtn = Str
					Exit Function
				End If
			Next
			DecodeCookie = StrRtn
		Else
			DecodeCookie = Str
		End If
	End Function
	Public Function BuildFile(ByVal sFile, ByVal sContent)
		On Error Resume Next
		Dim oFSO, oStream
		If CacheConfig(24) = "1" Then
			Set oStream = Server.CreateObject(CacheCompont(2))
			With oStream
				.Type = 2
				.Mode = 3
				.open
				'.Charset = "utf-8"
				.Charset = "gb2312"
				.Position = oStream.size
				.WriteText = sContent
				.SaveToFile sFile, 2
				.Close
			End With
			If Err.Number <> 0 Then
								'如果选用ADODB.Steam 则强制转换成Unicode
					If Right(LCase(sFile),4) <> ".xml" Then
						sContent = AnsiToUnicode(sContent)
					End if
					Set oStream = Server.CreateObject(CacheCompont(2))
					With oStream
						.Type = 2
						.Mode = 3
						.open
						'.Charset = "utf-8"
						.Charset = "gb2312"
						.Position = oStream.size
						.WriteText = sContent
						.SaveToFile sFile, 2
						.Close
					End With
					Err.Clear
				End If
			Set oStream = Nothing
		Else
			Set oFSO = Server.CreateObject(CacheCompont(1))
			Set oStream = oFSO.CreateTextFile(sFile,True)
			oStream.Write sContent
			oStream.Close
			'增加对特殊字符的保护，强制将内容转换成Unicode
			If Err.Number<>0 Then
				On Error Resume Next
				Set oStream = Server.CreateObject(CacheCompont(2))
				With oStream
					.Type = 2
					.Mode = 3
					.open
					'.Charset = "utf-8"
					.Charset = "gb2312"
					.Position = oStream.size
					.WriteText = sContent
					.SaveToFile sFile, 2
					.Close
				End With
				If Err.Number <> 0 Then
					sContent = AnsiToUnicode(sContent)
					Set oStream = Server.CreateObject(CacheCompont(2))
					With oStream
						.Type = 2
						.Mode = 3
						.open
						'.Charset = "utf-8"
						.Charset = "gb2312"
						.Position = oStream.size
						.WriteText = sContent
						.SaveToFile sFile, 2
						.Close
					End With
					Err.Clear
				End If
			End If
			Set oStream = Nothing
			Set oFSO = Nothing
		End If
	End Function
	'-----------Oblog4----------
	'sType:1-邀请码
	Public Function CheckOBCode(sCode, sType)
		Dim i, iAsc, rst,Sql
		sCode = UCase(Trim(sCode))
		CheckOBCode = False
		If Len(sCode)<>32 Then Exit FUnction
		For i = 1 To Len(sCode)
			iAsc = Asc(Mid(sCode, i, 1))
			'48~57,65~90
			If iAsc < 48 Or (iAsc > 57 And iAsc < 65) Or iAsc > 90 Then Exit Function
		Next
		If sType<>"" Then sType = CInt(sType)
		Sql="select iState From oblog_obcodes Where iState=0 And obcode='" & LCase(sCode) & "' "
		Sql =Sql & " "
		Set rst = Execute("select iState From oblog_obcodes Where iState=0 And obcode='" & LCase(sCode) & "' And iType=" & sType)
		If Not rst.EOF Then
			CheckOBCode = True
		End If
		rst.Close
		Set rst = Nothing
	End Function

	'检测用户发贴的许可
	Public Function CheckPostAccess()
		Dim rst,sql
		CheckPostAccess=""
		'首先进行新用户注册检验

		If CacheConfig(19)>0 Then
			If Int(datediff("n",l_uAddtime,Now))<Int(CacheConfig(19)) Then
				CheckPostAccess="系统设定您在注册后 " & CacheConfig(19) & " 分钟后才可以发布日志或者相册"
				Exit Function
			End If
		End If
		'检查每天最大的发帖数目
		If l_Group(10,0)<=0 Or l_Group(10,0)="" Then
			CheckPostAccess=""
		Else
			'此处也可加一个字段标记，本日该用户发布了多少篇日志
			sql = "select Count(logid) From oblog_log Where userid=" & l_uid & " And "
			If Is_Sqldata = 0 Then
				sql = sql & " Datediff('h',truetime,Now())<=24"
			Else
				sql = sql & " truetime BETWEEN DATEADD(Hour,-24,GETDATE()) AND GETDATE()"
			End if
			Set rst=Execute(sql)
			If rst(0)<l_Group(10,0) Then
				CheckPostAccess=""
			Else
				CheckPostAccess="您目前所属的组限制您24小时内只允许发布 " & l_Group(10,0) & " 篇日志<br/>您目前已经达到了该限额"
			End If
			Set rst=Nothing
		End If
	End Function

	'积分检查
	Public Function CheckScore(iScore)
		Dim rst
		CheckScore = False
'		If iScore >= 0 Then CheckScore = True: Exit Function
		Set rst = Execute("select scores From oblog_user Where userid=" & l_uId)
		If rst.EOF Then
			Set rst = Nothing
			Exit Function
		Else
			If rst(0) -  iScore > 0 Then
				CheckScore = True
			End If
		End If
		Set rst = Nothing
	End Function

	'给分,删分
	Public Function GiveScore(blogid, Score ,userid)
		Dim uid
		If userid<>"" Then
			uid = CLng(userid)
		Else
			uid = l_uId
		End if
		Score=CLng (Score)
		Execute ("Update oblog_user Set scores=scores+" & Score & " Where  userid=" & uid)
		If Score<0 Then Execute ("Update oblog_user Set scores=0 Where  userid=" & uid & " And  scores<0")
		If blogid <> "" Then
			Execute ("Update oblog_log Set scores=scores+" & Score & " Where logid=" & CLng (blogid) & "' And userid=" & uid)
		End If
	End Function

	'-------------------------------------------------------
	'内容保护模块!
	'-------------------------------------------------------
	'接管所有安全防护/内容过滤
	'内容类过滤,整合安全性过滤
	'关键字已经被分割成数组
	'此处的Content为返回参数
	Function CheckContent(byval Content, byval sType)
		Dim i,iCount,iLen,sKeep
		iCount=0
		Content=LCase(Content)
		'顶级过滤,直接封杀,系统对该用户进行计数,达到一定数目后,将该用户封禁
		For i=0 to Ubound(oblog.Keywords1)
			If Instr(Content,LCase(oblog.Keywords1(i)))>0 Then
	'				CheckContent=1 & "," & oblog.Keywords1(i)
				CheckContent=1
				Exit Function
			End If
		Next
		'次级过滤,提示审核
		For i=0 to Ubound(oblog.Keywords2)
			If Instr(Content,LCase(oblog.Keywords2(i)))>0 Then
				iCount=iCount+1
				sKeep= sKeep & "," & oblog.Keywords2(i)
				'If iCount>oblog.Setup(21) Then
				'	'此处借用了一个,
	'					CheckContent="2"& sKeep
					CheckContent=2
					Exit Function
				'End If
			End If
		Next
		'如果通过了第二次审核，则进入下一环节
		'一般过滤,全局字符替换
		For i=0 to Ubound(oblog.Keywords3)
			'如果是注册时存在，则直接跳出
			If sType="1" Then
				If Instr(Content,LCase(oblog.Keywords3(i)))>0 Then
					CheckContent=3
					Exit Function
				End If
			Else
			'如果是内容检测，则直接替换，不必执行查找过程
				Content=Replace(Content,oblog.Keywords3(i),"xxxx")
				CheckContent=3
			End If
		Next
		If CheckContent<>3 Then CheckContent=0
	End Function


	'注册时重复的用户名
	'注册禁止使用的用户名
	Function chk_regname(sUserName)
		Dim i
		chk_regname=0
		sUserName=Lcase(sUserName)
		'用户名不能为非英文字符
		If CacheConfig(6) <> "1" Then
			If chkDomain(sUserName)=false Then
					chk_regname=1
					Exit Function
			End If
		End if
		'用户名不能为系统禁止的关键字/审核字/过滤字
		If CheckContent(sUserName,1)<>0 Then
				chk_regname=2
				Exit Function
		End If
		'处理单独的注册关键字
		For i=0 to Ubound(oblog.Keywords4)
			If Trim (oblog.Keywords4(i))<>"" Then
				If Instr(sUserName,LCase(oblog.Keywords4(i)))>0 Then
					chk_regname=3
					Exit Function
				End If
			End if
		Next
		'如果不允许数字ID
		If en_nameisnum=0 Then
			If IsNumeric(sUserName) Then
				chk_regname=4
				Exit Function
			End if
		End if
		chk_regname=0
	End Function


	'进行IP控制
	Public Function ChkIpLock()
		If oblog.CheckAdmin(0) Then ChkIpLock = False :Exit Function
		Dim IPlock,i, sUserIP, sIP,BalckList,WhiteList,iCheck
		IPlock = False
		WhiteList = Application(Cache_Name & "_WhiteIp")
		BalckList = Application(Cache_Name & "_BlackIp")
		'如果无黑名单,则直接跳出
		If UBound(BalckList) < 0 Then
			ChkIpLock=False
			Exit Function
		End if
		'获取用户IP
		sUserIP = oblog.UserIp
		If sUserIP = "" Then Exit Function
		sUserIP = Split(UserIp, ".")

		If UBound(sUserIP) <> 3 Then Exit Function
		'检测白名单,白名单支持XXX.*.*.*,如果位于白名单内直接跳出检测流程
		For i = 0 To UBound(WhiteList)
			If WhiteList(i) <> "" Then
			  sIP = Split(WhiteList(i), ".")
			  If UBound(sIP) <> 3 Then Exit For
			  IPlock = false
			  If sUserIP(0) = sIP(0) Then
				If sUserIP(1) = sIP(1) Or  sIP(1)= "*" Then
					If sUserIP(2) = sIP(2) Or sIP(2)= "*" Then
						If sUserIP(3) = sIP(3) Or sIP(3)="*" Then
							ChkIpLock=false
							Exit Function
						End If
					End If
				End If
				End If
			End If
			Next
		'检测黑名单
		For i = 0 To UBound(BalckList)
			If BalckList(i) <> "" Then
				sIP = Split(BalckList(i), ".")
				If UBound(sIP) = 3  Then
					IPlock = True
					If (sUserIP(0) <> sIP(0)) And InStr(sIP(0), "*") = 0 Then IPlock = False
					If (sUserIP(1) <> sIP(1)) And InStr(sIP(1), "*") = 0 Then IPlock = False
					If (sUserIP(2) <> sIP(2)) And InStr(sIP(2), "*") = 0 Then IPlock = False
					If (sUserIP(3) <> sIP(3)) And InStr(sIP(3), "*") = 0 Then IPlock = False
					If IPlock Then Exit For
				End If
			End If
		Next
		ChkIpLock = IPlock
	End Function

	'进行白名单控制
	Public Function ChkWhiteIP(ByVal sUserIP)
		If oblog.CheckAdmin(0) Then ChkWhiteIP = True :Exit Function
		Dim IPlock,i, sIP,BalckList,WhiteList,iCheck
		ChkWhiteIP = False
		WhiteList = Application(Cache_Name & "_WhiteIp")
		'如果无黑名单,则直接跳出
		If UBound(WhiteList) < 0 Then
			Exit Function
		End if
		'获取用户IP
		sUserIP = oblog.UserIp
		If sUserIP = "" Then Exit Function
		sUserIP = Split(UserIp, ".")
		If UBound(sUserIP) <> 3 Then Exit Function
		'检测白名单,白名单支持XXX.*.*.*,如果位于白名单内直接跳出检测流程
		For i = 0 To UBound(WhiteList)
			If WhiteList(i) <> "" Then
			  sIP = Split(WhiteList(i), ".")
			  If UBound(sIP) <> 3 Then Exit For
			  IPlock = false
			  If sUserIP(0) = sIP(0) Then
				If sUserIP(1) = sIP(1) Or  sIP(1)= "*" Then
					If sUserIP(2) = sIP(2) Or sIP(2)= "*" Then
						If sUserIP(3) = sIP(3) Or sIP(3)="*" Then
							ChkWhiteIP=True
							Exit Function
						End If
					End If
				End If
				End If
			End If
		Next
	End Function

	'进行脚本过滤
	Function CheckScript(Content)
		Dim oRegExp,oMatch,spamCount
		Set oRegExp = New Regexp
		oRegExp.IgnoreCase = True
		oRegExp.Global = True
		oRegExp.pattern ="<script.+?/script>"
		Content=oRegExp.replace(Content,"")
		Set oRegExp=Nothing
	End Function

	'进行多媒体对象检测
	'提取媒体文件,清理播放器
	Function CheckMedia(Content)
		Dim oRegExp,oRegExp1,oMatch,Matches,oMatch1,Matches1
		Dim sFiles1,sFiles2,sFile
		sFiles="swf,mp3,rm,ram,rmvb,mp4,wma,wav,avi"
		Set oRegExp = New Regexp
		oRegExp.IgnoreCase = True
		oRegExp.Global = True
		Set oRegExp1 = New Regexp
		oRegExp1.IgnoreCase = True
		oRegExp1.Global = True

		'媒体文件
		oRegExp.pattern ="<object.+?>"
		Set Matches=oRegExp.Execute(Content)
		For Each oMatch In Matches
			oRegExp1.pattern="http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?"
			Set Matches1=oRegExp.Execute(oMatch.Value)
			For Each oMathch1 In Matches1
				'只取媒体文件
				sFile=Split(oMathch1.value,".")
				If InStr(sFiles1,sFile(UBound(sFile)))>0 Then
					strFiles2="<a href=""" &  oMathch1.value & """ target=""_blank"">" & oMathch1.value & "</a><br>"
				End If
			Next
		Next
		'清空
		oRegExp.pattern ="<object.+?/object>"
		Content=oRegExp1.replace(Content,"")
		oRegExp.pattern ="<em.+?>"
		Set Matches=oRegExp.Execute(Content)
		For Each oMatch In Matches
			oRegExp1.pattern="http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?"
			Set Matches1=oRegExp.Execute(oMatch.Value)
			For Each oMathch1 In Matches1
				'只取媒体文件
				sFile=Split(oMathch1.value,".")
				If InStr(sFiles1,sFile(UBound(sFile)))>0 Then
					strFiles2="<a href=""" &  oMathch1.value & """ target=""_blank"">" & oMathch1.value & "</a><br>"
				End If
			Next
		Next
		oRegExp.pattern ="<em.+?/em>"
		Content=oRegExp1.replace(Content,"")
		Set oRegExp1=othing
		Set oRegExp2=othing
	End Function

	Function ubb_comment(strContent)
		Dim re

		If IsNull(strContent) THen
			ubb_comment=""
			Exit Function
		End If

		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True
		'以下过滤html代码
		strContent = Replace(strContent, "<br />", "[br]")
		strContent = Replace(strContent, ">", "&gt;")
		strContent = Replace(strContent, "<", "&lt;")
		strContent = Replace(strContent, Chr(32), " ")
		strContent = Replace(strContent, Chr(9), " ")
		strContent = Replace(strContent, Chr(34), "&quot;")
		'strContent = Replace(strContent, CHR(39), "&#39;")
		strContent = Replace(strContent, Chr(13), "")
		strContent = Replace(strContent, Chr(10), "<br /> ")
		strContent = Replace(strContent, "[br]", "<br />")
		'以下过滤ubb标签
		re.Pattern="(\[EMOT\])(.[^\[]*)(\[\/EMOT\])"
		strContent= re.replace(strContent,"<img src="""&blogurl&"editor/images/emot/face"&"$2"&".gif"&""" />")
		re.Pattern="\[i\](.[^\[]*)(\[\/i\])"
		strContent=re.replace(strContent,"<em>$1</em>")
		re.Pattern="\[u\](.[^\[]*)(\[\/u\])"
		strContent=re.replace(strContent,"<u>$1</u>")
		re.Pattern="\[b\](.[^\[]*)(\[\/b\])"
		strContent=re.replace(strContent,"<strong>$1</strong>")
'		re.Pattern="\[QUOTE\](.[^\[]*)(\[\/QUOTE\])"
'		strContent=re.replace(strContent,"<div class=""quote"">$1</div><br>")
		re.Pattern="\[QUOTE\]"
		strContent=re.replace(strContent,"<div class=""quote"">")
		re.Pattern="\[\/QUOTE\]"
		strContent=re.replace(strContent,"</div>")
		Set re=Nothing
		ubb_comment=strContent
	End Function
	'载入编辑器，stype值为1可上传,0不可上传
	Sub MakeEditorText(sInput,stype,width,height)
		If l_isUbb > 0 Then C_Editor_Type = l_isUbb
		If C_Editor_Type = 2 Then Exit Sub
		If sInput = "" Then sInput = "edit"
		Select Case C_Editor_Type
			Case 1
	%>
			<script language=JavaScript src="<%=C_Editor%>/scripts/language/schi/editor_lang.js"></script>
			<script language=JavaScript src="<%=C_Editor%>/scripts/innovaeditor.js"></script>
			<script language="JavaScript">
				var oEdit1 = new InnovaEditor("oEdit1");

				//STEP 2: Asset Manager Localization: Add querystring lang=english/danish/dutch...
				//oEdit1.css="/editor/scripts/style/editor.css"
			<%if oblog.CacheConfig(53) = "1" Then%>
				oEdit1.mode="XHTMLBody";
			<%end if%>
				oEdit1.width=<%=width%>;
				oEdit1.height=<%=height%>;
				oEdit1.cmdCustomObject = "modelessDialogShow('<%=blogdir%>editor/scripts/emot.htm',280,200)";
			<%if stype = 1 Then %>
				oEdit1.cmdAssetManager="modalDialogShow('<%=blogdir%>editupload.asp',640,465)";
			<%End If%>
			<%if oblog.CacheConfig(53) = "1" Then%>
				oEdit1.btnHTMLSource=false;
				oEdit1.btnXHTMLSource=true;
			<%end if%>
				oEdit1.REPLACE("<%=sInput%>");
				oEdit1.focus();
			</script>
	<%
			Case 2
		End Select
		%>
	<%
	End Sub
	 '发送系统信息
	Sub SendSysMsg(fromId,toId,toName,toContent)

	End Sub

	'CheckAdmin系统管理员1,内容管理员2,任意管理员0
	Public Function CheckAdmin(sType)
		Dim admin_name,admin_password,sql,rs
		CheckAdmin=False
		admin_name=filt_badstr(session("adminname"))
		admin_password=filt_badstr(session("adminpassword"))
		If admin_name = "" Or admin_password = "" Then
'			If sType <> 1 Then
				admin_name=filt_badstr(session("m_name"))
				admin_password=filt_badstr(session("m_pwd"))
'			End If
		End if
'		If IsEmpty(admin_name) Or admin_name="" Then Exit Function
		sql="select id,password,roleid from oblog_admin where username='" & admin_name & "' and password='"&admin_password&"'"
		If Not IsObject(conn) Then link_database
		Set rs=conn.execute(sql)
		if Not rs.eof Then
			If sType = 1 Then
				If rs(2) <> 0 Then Exit Function
			ElseIf sType = 2 Then
				If rs(2) = -1 Then Exit Function
			End if
			If rs(1)=admin_password Then
				rs.close
				set rs=nothing
				CheckAdmin=True
				Exit Function
			End If
		End if
		rs.close
		Set rs=Nothing
	End Function

	'验证用户提交的域名根是否合法
	Public Function CheckDomainRoot(R_DomainRoot,sType)
		CheckDomainRoot=False
		Dim DomainRoot,i
		If sType = 0 Then
			DomainRoot=Trim(CacheConfig(4))
		ElseIf sType = 1 Then
			DomainRoot=Trim(CacheConfig(75))
		End if
		R_DomainRoot=Trim (R_DomainRoot)
		If DomainRoot="" Or CacheConfig(5) = 0 Then Exit Function
		If InStr(DomainRoot,"|")<0 Then
			If R_DomainRoot=DomainRoot Then
				CheckDomainRoot=True
				Exit Function
			End If
		Else
			DomainRoot=Split(DomainRoot,"|")
			For i=0 To UBound(DomainRoot)
				If R_DomainRoot = DomainRoot(i) Then
					CheckDomainRoot=True
					Exit Function
				End If
			Next
		End if
	End Function

	'过滤掉flash UBB标记
	Function FilterUBBFlash(byval strFlash)
		Dim strFlash1,t
		t=0
		strFlash1=LCase(strFlash)
		If InStr(strFlash1,"[/flash]")>0 Then
			strFlash1 = Replace(strFlash1,"[/flash]","[ /flash ]")
			strFlash1 = Replace(strFlash1,"[flash","[ flash ")
			t=1
		end if
		if InStr(strFlash1,"[/mp]")>0 Then
			strFlash1 = Replace(strFlash1,"[/mp]","[ /mp ]")
			strFlash1 = Replace(strFlash1,"[mp","[ mp ")
			t=1
		end if
		if InStr(strFlash1,"[/rm]")>0 Then
			strFlash1 = Replace(strFlash1,"[/rm]","[ /rm ]")
			strFlash1 = Replace(strFlash1,"[rm","[ rm ")
			t=1
		End If
		if InStr(strFlash1,"[/url]")>0 Then
			strFlash1 = Replace(strFlash1,"[/url]","[ /url ]")
			strFlash1 = Replace(strFlash1,"[url","[ url ")
			t=1
		End If
		if InStr(strFlash1,"meta")>0 Then
			strFlash1 = Replace(strFlash1,"meta","ｍeta")
			t=1
		End If
		if InStr(strFlash1,"embed")>0 Then
			strFlash1 = Replace(strFlash1,"embed","ｅmbed")
			t=1
		End If
		if t=1 then
			FilterUBBFlash=strFlash1
		else
			FilterUBBFlash=strFlash
		end if
	End Function

	'封IP
	Public Sub KillIP(sIP)
		'如果在白名单则不进行锁定IP操作
		If ChkWhiteIP(sIP) Then Exit Sub
		Dim rstCache
		Set rstCache = Server.CreateObject("Adodb.RecordSet")
		rstCache.Open "select * From  oblog_config Where id=5",conn,1,3
		If InStr(rstCache("ob_value"),sIP)<=0 Then
			rstCache("ob_value")=FilterEmpty(rstCache("ob_value")& vbCrLf & sIP)
			rstCache.Update
		End If
		rstCache.Close
		Set rstCache=Nothing
		ReLoadCache
	End Sub
	'过滤关键字、黑白名单ip中的空行
	Function FilterEmpty(badstr)
		Dim arrStr,strReturn,i
		badstr=Trim (badstr)
		If badstr= "" Then
			FilterEmpty=badstr
			Exit Function
		End if
		If InStr (badstr,vbcrlf)>0 Then
			arrStr = Split (badstr,vbcrlf)
			For i = 0 To UBound(arrStr)
				If arrStr(i)<>"" Then
					strReturn = strReturn & vbcrlf & arrStr(i)
				End if
			Next
			strReturn = Replace (strReturn,vbcrlf,"",1,1,0)
		Else
			strReturn = badstr
		End If
		FilterEmpty = strReturn
	End Function
	'统计日志数目
	'sType为"+"或者"-"
	Sub log_count(ByVal userID,ByVal logid,ByVal subjectID,ByVal classID,ByVal sType)
        Execute ("UPDATE oblog_user SET log_count = log_count"&sType&"1 WHERE userid=" & CLng (userID))
		Execute ("UPDATE [oblog_setup] SET log_count = log_count"&sType&"1")
		If logid <>"" Then
			If sType = "+" Then
				Execute ("UPDATE oblog_comment SET isdel = 1 WHERE mainid=" & CLng (logid))
			ElseIf sType = "-" Then
				Execute ("UPDATE oblog_comment SET isdel = 0 WHERE mainid=" & CLng (logid))
			End If
		End if
		If subjectID <> "" Then
			Execute ("UPDATE oblog_subject SET subjectlognum = subjectlognum"&sType&"1 WHERE subjectid = " & CLng (subjectID))
		End If
		If classID<>"" Then
			Execute ("UPDATE [oblog_logclass] SET classlognum = classlognum"&sType&"1 WHERE id = " & CLng (classID))
		End IF
	End Sub

	'-------------------------------------------------------
	'分类缓存节
	'-------------------------------------------------------
	Private bUpdateClass,bUpdateSysSkin
	'是否开启缓存分类的调试模式
	Private Property Get Cache_Debug_Mode
		Cache_Debug_Mode = False
	End Property
	Public Sub ResetClassCache()
		bUpdateClass=OB_IIF(Application(Cache_Name & "_Class_NeedUpdate"),True)
		bUpdateSysSkin=Application(Cache_Name & "_SysSkin_NeedUpdate")
		If bUpdateSysSkin="" Then bUpdateSysSkin=True
		If Cache_Debug_Mode Then bUpdateClass = True
		If bUpdateClass Then
			Call ClassArray(1,0)
			Call ClassArray(2,0)
			Call ClassArray(2,1)
			Call ClassArray(2,2)
			Call ClassString(1,0)
			Call ClassString(2,0)
			Call ClassString(2,1)
			Call ClassString(2,2)
			Application.Unlock
			Application(Cache_Name & "_Class_NeedUpdate")=false
			Application.lock
		End If
	End Sub
	'获取分类数组
	'sType1:1-用户分类;2-日志分类
	'sType2:0-日志;1-相册;2-群组分类
	Function ClassArray(ByVal sType1,ByVal sType2)
		Dim sqlClass,SqlStr
		Dim rst,rst1
		Dim thisArr,ArrayFields
		ReDim ArrayFields(4)
		ArrayFields(0) = "id"
		ArrayFields(1) = "classname"
		ArrayFields(2) = "depth"
		ArrayFields(3) = "NextId"
		ArrayFields(4) = "ParentPath"
		SqlStr = Join(ArrayFields,",")
		If Cache_Debug_Mode Then bUpdateClass = True
		If bUpdateClass Then
			Set rst=Server.CreateObject("Adodb.Recordset")
			rst.CursorLocation=3
			If sType1 = 1 Then
				sqlClass = "select "&SqlStr&" From oblog_userclass order by RootID,OrderID"
			Else
				sqlClass = "select "&SqlStr&" From oblog_logclass  Where idType=" & sType2 & " order by RootID,OrderID"
			End If
			Set rst=Execute(sqlClass)
'			rst.Open SqlClass,conn,1,1
			If Not rst.Eof Then
				ThisArr=rst.GetRows(-1,0,ArrayFields)
			End if
			rst.Close
			Set rst=Nothing
			Application.unLock
			Application(Cache_Name & "_Class_Arr_"& sType1 & "_" & sType2)=ThisArr
			Application.Lock
		End If
		ClassArray=Application(Cache_Name & "_Class_Arr_"& sType1 & "_" & sType2)

		''DEBUG MODE
		If Cache_Debug_Mode Then
			Dim iRecFirst,iRecLast,iFieldFirst,iFieldLast,arrDBData,i,j
			arrDBData=ClassArray
			iRecFirst   = LBound(arrDBData, 2)
			iRecLast    = UBound(arrDBData, 2)
			iFieldFirst = LBound(arrDBData, 1)
			iFieldLast  = UBound(arrDBData, 1)
			' Loop through the records (second dimension of the array)
			Response.Write "<table border = 1>"
			For I = iRecFirst To iRecLast
				' A table row for each record
				Response.Write "<tr>" & vbCrLf

				' Loop through the fields (first dimension of the array)
				For J = iFieldFirst To iFieldLast
					' A table cell for each field
					Response.Write vbTab & "<td>(" & j & "," & i& "):" & arrDBData(J, I) & "</td>" & vbCrLf
				Next ' J

				Response.Write "</tr>" & vbCrLf
			Next ' I
			Response.Write "</table>"
			Response.Write "<hr>"
		End If
	End Function

	'获取分类Select中的字串
	'sType1:1-用户分类;2-日志分类
	'sType2:0-日志;1-相册
	Public Function ClassString(byval sType1,byval sType2)
		Dim rst, sqlClass, sTmp, tmpDepth, i,j,thisArr
		Dim arrShowLine(20),sRet
		For i = 0 To UBound(arrShowLine)
			arrShowLine(i) = False
		Next
		'If bUpdateClass=false Then
		'	ClassString=Application(Cache_Name & "_Class_String_"& sType1 & "_" & sType2)
		'	Exit Function
		'End If
		sRet = "<option value='0'>请选择类别</option>"
		'Response.Write Typename(Application(Cache_Name & "_Class_Rst_"& sType1 & "_" & sType2))
		'Response.End
		thisArr=Application(Cache_Name & "_Class_Arr_"& sType1 & "_" & sType2)
		If IsArray(thisArr) Then
			For j=0 To UBound(thisArr,2)
				tmpDepth = thisArr(2,j)
				If thisArr(3,j) > 0 Then
					arrShowLine(tmpDepth) = True
				Else
					arrShowLine(tmpDepth) = False
				End If
					sTmp = "<option value='" & thisArr(0,j) & "'>"

				If tmpDepth > 0 Then
					For i = 1 To tmpDepth
						sTmp = sTmp & "&nbsp;&nbsp;"
						If i = tmpDepth Then
							If thisArr(3,j) > 0 Then
								sTmp = sTmp & "├&nbsp;"
							Else
								sTmp = sTmp & "└&nbsp;"
							End If
						Else
							If arrShowLine(i) = True Then
								sTmp = sTmp & "│"
							Else
								sTmp = sTmp & "&nbsp;"
							End If
						End If
					Next
				End If
				sTmp = sTmp & thisArr(1,j)
				sTmp = sTmp & "</option>"
				sRet= sRet & sTmp
			Next
		End if
		Application.Unlock
		Application(Cache_Name & "_Class_String_"& sType1 & "_" & sType2)=sRet
		Application.lock
		ClassString=Application(Cache_Name & "_Class_String_"& sType1 & "_" & sType2)
		sRet=""
	End Function

	'日志、用户资料编辑时Select控件显示
	'动态加载,仅作替换
	Public Function SelectedClassString(byval sType1,byval sType2,byval sSelected)
		Dim sClass
		sClass=ClassString(sType1,sType2)
		If Int(sSelected) > 0  Then
			'<option value='" & rst("id") & "'>
			sClass=Replace(sClass,"<option value='" & sSelected & "'>","<option value='" & sSelected & "' Selected>")
		End If
		SelectedClassString=sClass
		sClass=""
	End Function

	'获取单一分类的名称
	'sType1:1-OBLOG_USERCLASS表;2-OBLOG_LOGCLASS表
	'sType2:0-日志;1-相册;2-群组分类
	'sClassId:当前选中的分类ID
	Public Function GetClassName(Byval sType1,Byval sType2,sClassId)
		Dim thisArr,i
		thisArr=ClassArray(sType1,sType2)
'		OB_DEBUG sType2,1
		For i=0 to UBound(thisArr,2)
			If sClassId=thisArr(0,i) Then
				GetClassName=thisArr(1,i)
				Exit Function
			End If
		Next
		If IsNull(GetClassName) Or GetClassName = 0 Then GetClassName = "无分类"
	End Function

	'获取用户可发布的日志与相册分类
	'系统只控制到第一级别,需取出当前级别及其子级别
	Public Function UserPostClass(byval sType1,byval sType2,CurrentID)
		Dim rsClass, sqlClass, sTmp, tmpDepth, i,j,Sql,thisArr,sRet
		Dim arrShowLine(20)
		For i = 0 To UBound(arrShowLine)
			arrShowLine(i) = False
		Next
		'处理类别
		Dim sClass,sClass1,aClass,show_Postclass
		sClass=Trim(oblog.l_Group(9,0))
		If sClass="" Or IsNull(sClass) Then
			 '取总分类
'			 UserPostClass=ClassString(sType1,sType2)
			 UserPostClass=SelectedClassString(sType1,sType2,CurrentID)
			 Exit Function
		End If
		thisArr=ClassArray(sType1,sType2)
		sClass="," & sClass & ","
		sRet = "<option value='0'>请选择类别</option>"
		For i=0 To UBound(thisArr,2)
				'获取该类别的父类别，如果本身是父类别则默认为其自己
				If OB_IIF(thisArr(4,i),0)="0" Then
					sClass1=thisArr(0,i)
				Else
					aClass=Split(thisArr(4,i),",")
					sClass1=aClass(1)
				End If
				'Response.Write sClass1 & "<br>"
				If InStr(sClass,"," & sClass1 & ",") Then
					tmpDepth = thisArr(2,i)
					'Response.Write tmpDepth & "<br>"
					If thisArr(3,i) > 0 Then
						arrShowLine(tmpDepth) = True
					Else
						arrShowLine(tmpDepth) = False
					End If
					sTmp = "<option value='" & thisArr(0,i) & "'"
					If CurrentID > 0 And thisArr(0,i) = CurrentID Then
						 sTmp = sTmp & " selected"
					End If
					sTmp = sTmp & ">"

					If tmpDepth > 0 Then
						For j = 1 To tmpDepth
							'Response.Write "yy" & "<br>"
							sTmp = sTmp & "&nbsp;&nbsp;"
							If j = tmpDepth Then
								sTmp = sTmp & "├&nbsp;"
							Else
								If arrShowLine(j) = True Then
									sTmp = sTmp & "│"
								Else
									sTmp = sTmp & "&nbsp;"
								End If
							End If
						Next
					End If
					sTmp = sTmp & thisArr(1,i)
					sTmp = sTmp & "</option>"
					sRet = sRet & sTmp
				End If
			Next
			UserPostClass=sRet
	End Function

	Sub ClearOldOBCodes()
		Execute("Delete From oblog_obcodes Where istate=0 And datediff("&G_Sql_d&",createtime,"&G_Sql_Now&")>=15")
	End Sub
	'清理旧的用户回收站日志
	Sub ClearOldUserRLog()
		Dim deltime
		On Error Resume Next
		deltime = int(oblog.CacheConfig(87))
		If Err Then Err.clear '防护第一次未保存设置的时候无法进入系统后台
		If int(deltime) < 60 Then deltime = 60
		If Err Then Err.clear:deltime = 60 '防护未保存设置的时候为空的时候误删除小于60天的日志
		Execute("Delete From oblog_log Where isdel=1 And datediff("&G_Sql_d&",truetime,"&G_Sql_Now&")>="&int(deltime)&"")
	End Sub

	Public Sub CountGroupUser()
		Dim rs,rst
		Set rs=Server.CreateObject("Adodb.Recordset")
		Set rst=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select * From oblog_groups",conn,1,3
		'重新进行计数
		rst.Open "Select Count(UserId),user_group From oblog_user Where user_group>0 Group By user_group",conn,1,3
		Do While Not rs.Eof
			rst.Filter="user_group=" & rs("groupid")
			If rst.Eof Then
				Execute("Update oblog_groups Set g_members=0 Where groupid=" & rs("groupid"))
			Else
				Execute("Update oblog_groups Set g_members=" & rst(0) & " Where groupid=" & rs("groupid"))
			End If
			rs.MoveNext
		Loop
		Set rs=Nothing
		Set rst=Nothing
	End Sub

	'userid用户id 可为数组或者1,2,3
	'arrayUBound 指定传入数组的Ubound
	'生成JS函数的唯一标识，可任意，但是必须保证唯一性
	Public Function GetNickNameById(ByVal userid,ByVal arrayUBound ,ByVal Strings)
		On Error Resume Next
		Dim arrayList,RS,i,name,arrayListTemp,arrayListTempUserId
		Dim strTemp,showTemp,strTempUserId
		Dim userTemp
		i = 0
		ReDim arrayListTemp(arrayUBound-1)
		ReDim arrayListTempUserid(arrayUBound-1)
		If Not IsArray (userid) Then
			userid = FilterIDs(userid)
			If userid = "" Then Exit Function
			userTemp = userid
		Else
			arrayList = userid
			userTemp = Join(arrayList,",")
			userTemp = FilterIDs(userTemp)
		End If
		If userTemp = "" Then Exit Function
		Set RS = Execute ("SELECT username,nickname,userid FROM oblog_user WHERE userid IN ("&userTemp&")")
		Do While Not RS.Eof
			arrayListTemp(i) = "'"&Replace(OB_IIF(RS(1),RS(0)),"'","‘")&"'"
			arrayListTempUserid(i) = "'"&RS(2)&"'"
			i = i + 1
			RS.MoveNext
		Loop
		strTemp = Join(arrayListTemp,",")
		strTemp = FilterStrings(strTemp)
		strTempUserId = Join(arrayListTempUserId,",")
		strTempUserId = FilterStrings(strTempUserId)
		showTemp = vbcrlf & "<script language=""JavaScript"">"& vbcrlf
		showTemp = showTemp &"var arrayList_"&Strings&" = new Array (["&strTemp&"],["&strTempUserId&"]);"& vbcrlf
		showTemp = showTemp &"for(var i = 0 ;i<arrayList_"&Strings&"[0].length;i++)"& vbcrlf
		showTemp = showTemp &"	{ var obj=document.getElementsByName('nickname_'+arrayList_"&Strings&"[1][i]);"& vbcrlf
		showTemp = showTemp &"		if (obj)"& vbcrlf
		showTemp = showTemp &"		{ for (var j=0;j<obj.length;j++)"& vbcrlf
		showTemp = showTemp &"			{"& vbcrlf
		showTemp = showTemp &"				obj[j].innerHTML=arrayList_"&Strings&"[0][i];"& vbcrlf
		showTemp = showTemp &"			}"& vbcrlf
		showTemp = showTemp &"		}"& vbcrlf
		showTemp = showTemp &"	}"& vbcrlf
		showTemp = showTemp & "</script>"& vbcrlf
		Set RS = Nothing

		GetNickNameById = showTemp
	End Function
	'作用同GetNickNameById，处理方式不同，此处是做字符串替换
	Public Function GetNameNameByUserId(ByVal userid,ByVal strings)
		Dim arrayList,arrayListTemp
		Dim RS,idS
		Dim showString,allNickName
		Dim i,tuid
		If Not IsArray (userid) Then
			userid = FilterIDs(userid)
			If userid = "" Then Exit Function
			IDS = userid
		Else
			arrayList = userid
			IDS = Join(arrayList,",")
			IDS = FilterIDs(IDS)
		End If
		Set RS = oblog.Execute ("SELECT username ,nickname,userid FROM oblog_user WHERE userid IN ("&IDS&")")
		While Not RS.eof
			allNickName=allNickName&RS(2)&"!!??(("&OB_IIF(RS(1),RS(0))&"##))=="
			RS.MoveNext
		Wend
		arrayListTemp=Split(allNickName,"##))==")
		'循环数组
		For i=0 To UBound(arrayListTemp)
			'取userid
			tuid=Split(arrayListTemp(i),"!!??((")(0)
			'替换昵称
			showString=Replace(strings,"nickname_"&tuid,GetsubName(tuid,allNickName))
			i=i+1
		Next
		Set RS = Nothing
		GetNameNameByUserId = showString
	End Function

	'-------------------------
	Public 	Sub reset_album_cover(ByVal  uid,ByVal idd)
		On Error Resume Next
		If uid="" Or IsNull(uid) Then Exit Sub
		Dim rst,rsu,rsp
	set rst=Server.CreateObject("adodb.recordset")
	rst.open "select subjectid,subjectlognum,subjecttype,photo_path from oblog_subject WHERE subjecttype = 1  AND userid="&uid,conn,2,2
	while not rst.eof
		Set rsu=oblog.Execute("SELECT TOP 1 fileid,photo_path,(SELECT COUNT(photoid) FROM oblog_album WHERE (ishide = 0 OR ishide IS NULL) AND userclassid = "&rst(0)&") AS pnum FROM oblog_Album WHERE (isHide = 0 OR isHide IS NULL) AND (userClassId = "&rst(0)&") order by Is_Album_default desc,photoid desc")
		if not rsu.eof Then
		rst("photo_path")=rsu(1)
		rst("subjectlognum")=rsu(2)
		If idd<>"" And idd<> 0 And Int(rsu(0))=Int(idd) Then oblog.execute("update oblog_Album set is_album_default = 0 where userClassId = "&rst(0))
		else
		rst("photo_path")=""
		rst("subjectlognum")=0
		End If
		Set rsu=Nothing
		rst.update
		rst.movenext
	wend
	rst.close
	Set rst = Nothing

	End Sub
End Class

Class AjaxXml
	Private m_contentType,m_encoding,m_xml

	Private Sub Class_Initialize()
		m_contentType = "text/xml"
		m_encoding = "gb2312"
		m_xml=true
	End sub

	Public sub re(result)
		Response.contentType = m_contentType
		Response.Expires=0
		Response.Clear
		Response.Write serialize(result)
	End Sub

	Private function serialize(result)
		Dim restr,i
		if m_xml then
			restr = "<?xml version=""1.0"" encoding="""&m_encoding&"""?>"
			restr = restr+"<Response>"
			if IsArray(result) then
				For i=0 to UBound(result)
					restr = restr + "<item><![CDATA["&result(i)&"]]></item>"
				next
			else
				restr = restr + result
			end If
			restr = restr + "</Response>"
		else
			restr = result
		end if
		serialize = restr
	end Function

End Class
%>
