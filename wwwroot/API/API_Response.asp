<%@ LANGUAGE = VBScript CodePage = 936%>
<!-- #include file="../Conn.asp" -->
<!-- #include file="../inc/class_sys.asp" -->
<!-- #include file="Class_API.asp" -->
<!-- #include file="../inc/md5.asp" -->
<!--#include file="../inc/Cls_XmlDoc.asp"-->
<%
Dim FoundErr,ErrMsg
Dim Action,syskey,username,password,CookieDate,appid
Dim Sex,QQ,MSN,UserStatus,TrueName,Birthday,TelePhone,HomePage,userip,email,Question,Answer,province,city,address
Dim oblog,XMLdom,blogAPI
set oblog=new class_sys
oblog.start
If Request.QueryString("syskey")<>"" Then
	syskey=LCase(Request.QueryString("syskey"))
	username=oblog.filt_badstr(Trim(Request("username")))
	If ChkSyskey Then
		Dim TruePassWord
		TruePassWord = RndPassword(16)
		If Request.QueryString("password")<>"" Then
			password=oblog.filt_badstr(Request("password"))
			CookieDate=Trim(Request("savecookie"))
			If CookieDate="0" Or CookieDate="" Then CookieDate="1"
			oblog.Execute ("UPDATE oblog_user SET TruePassWord = '"&TruePassWord&"' WHERE username = '"&UserName&"' AND password = '"&password&"'")
			oblog.savecookie UserName,TruePassWord,CookieDate
		Else
			Call LogoutUser()
		End If
	End If
Else
	Set blogAPI = New DPO_API_OBLOG
	blogAPI.LoadXmlFile False
	Set XMLdom = Server.CreateObject("Microsoft.XMLDOM")
	XMLdom.Async = False
	XMLdom.Load(Request)
	If API_Enable=False Then
		ErrMsg=("系统并未开启整合接口！")
		FoundErr=True
		blogAPI.SendResult 1, ErrMsg
		Set blogAPI=Nothing
		Response.End
	End If
	If XMLdom.parseError.errorCode <> 0 Then
		ErrMsg=("接收数据出错，请重试！")
		FoundErr=True
		blogAPI.SendResult 1, ErrMsg
		Set blogAPI=Nothing
		Response.End
	Else
		appid = XMLdom.documentElement.selectSingleNode("//appid").text
		syskey = XMLdom.documentElement.selectSingleNode("//syskey").text
		Action = XMLdom.documentElement.selectSingleNode("//action").text
		UserName=oblog.filt_badstr (XMLdom.documentElement.selectSingleNode("//username").text)
	End If
	If ChkSyskey Then
		select Case Action
			Case "reguser"
				Call reguser()
			Case "login"
				Call ot_chklogin (UserName,PassWord,CookieDate)
			Case "logout"
				Call LogoutUser()
			Case "update"
				Call ModifyUserInfo()
			Case "delete"
				Call DelUser()
			Case "getinfo"
				Call getuserinfo()
			Case "checkname"
				Call Checkname()
		End select
	End If
	'将处理结果返回到类文件中处理以便提交方得到该信息
	If FoundErr Then
		blogAPI.SendResult 1, ErrMsg
	Else
		blogAPI.SendResult 0,""
	End If
	Set XMLdom=Nothing
	Set blogAPI=Nothing
End If
Set oblog=Nothing

Sub Checkname()
	Dim chk_regname
	chk_regname=oblog.chk_regname(UserName)
	EMail=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//email").text)
	if oblog.CacheConfig(15) = 0 Then
		ErrMsg=ErrMsg&"当前系统已关闭注册！"
		FoundErr=True
		Exit Sub
	End If
	If oblog.chkiplock() Then
		ErrMsg=ErrMsg&"对不起，你的IP已被锁定,不允许注册！"
		FoundErr=True
		Exit Sub
	End If
	if UserName="" Then
		ErrMsg=ErrMsg&("用户名不允许为空！")
		FoundErr=True
	End If
	if chk_regname>0 then
'		if chk_regname = 1 Then ErrMsg=ErrMsg&("用户名不合规范，只能使用小写字母，数字及下划线！")
		if chk_regname = 2 Then ErrMsg=ErrMsg&("用户名中含有系统不允许的字符！")
		if chk_regname = 3 Then ErrMsg=ErrMsg&("用户名中含有系统保留注册的字符！")
		if chk_regname = 4 Then ErrMsg=ErrMsg&("用户名中不允许全部为数字！")
		If ErrMsg<>"" Then FoundErr=True
	End If
	Dim rstc
	Set rstc=oblog.execute ("select * from oblog_user where username='"&UserName&"'")
	If Not rstc.eof Then
		ErrMsg=ErrMsg&("用户名已经存在，请更换！")
		FoundErr=True
	End If
	rstc.close
	Set rstc=Nothing
End Sub
'oblog用户整合的注册函数
Sub reguser()
	Dim chk_regname
	chk_regname=oblog.chk_regname(UserName)
	Call GetXML()
	if oblog.CacheConfig(15) = 0 Then
		ErrMsg="当前系统已关闭注册！"
		FoundErr=True
		Exit Sub
	End If
	If oblog.chkiplock() Then
		ErrMsg="对不起，你的IP已被锁定,不允许注册！"
		FoundErr=True
		Exit Sub
	End If
	if chk_regname>0 then
'		if chk_regname = 1 Then ErrMsg=ErrMsg&("用户名不合规范，只能使用小写字母，数字及下划线！")
		if chk_regname = 2 Then ErrMsg=ErrMsg&("用户名中含有系统不允许的字符！")
		if chk_regname = 3 Then ErrMsg=ErrMsg&("用户名中含有系统保留注册的字符！")
		if chk_regname = 4 Then ErrMsg=ErrMsg&("用户名中不允许全部为数字！")
		If ErrMsg<>"" Then FoundErr=True
	End If
	If PassWord="" Then
		ErrMsg=ErrMsg&("密码不能为空！")
		FoundErr=True
	End If
'	If Question="" Then
'		ErrMsg=ErrMsg&("提示问题不能为空！")
'		FoundErr=True
'	End If
'	If Answer="" Then
'		ErrMsg=ErrMsg&("提示答案不能为空！")
'		FoundErr=True
'	End If
	If EMail="" Then
		ErrMsg=ErrMsg&("EMail不能为空！")
		FoundErr=True
	End If
	If oblog.CacheConfig(22) = 1 Then
	If Not onlyEMail(EMail) Then
		ErrMsg=ErrMsg&("EMail不能重复！")
		FoundErr=True
	End If
	End If
	If FoundErr=True Then Exit Sub
	Dim Reguserlevel
	if oblog.CacheConfig(18) = 1 Then reguserlevel=6 else reguserlevel=7
	Dim rsreg
	if Not IsObject(conn) Then link_database
	Set rsreg=Server.CreateObject("adodb.recordset")
	rsreg.open "select * from [oblog_user] where UserName='"& oblog.filt_badstr(UserName) &"'",conn,1,3
	If rsreg.eof Then
		rsreg.addnew
		rsreg("UserName")=UserName
		rsreg("PassWord")=md5(PassWord)
'		rsreg("Question")=Question
'		rsreg("Answer")=md5(Answer)
		rsreg("userEMail")=EMail
		rsreg("user_level")=reguserlevel
		rsreg("blogname")=UserName & "的blog"
		rsreg("user_isbest")=0
		rsreg("province")=province
		rsreg("city")=city
		If oblog.chkdomain(UserName)=False Then
			rsreg("Nickname")=UserName
		End If
		rsreg("adddate")=oblog.ServerDate(Now())
        rsreg("regip") = oblog.userip
        rsreg("lastloginip") = oblog.userip
		rsreg("lastlogintime")=oblog.ServerDate(Now())
		rsreg("user_dir") =oblog.setup(8,0)
        rsreg("user_group") = oblog.defaultGroup
        rsreg("scores") = oblog.cacheScores(1)
        rsreg("newbie") = 1
		rsreg.update
		oblog.execute("update oblog_setup set user_count=user_count+1")
		oblog.execute("update oblog_user set user_folder=userid where UserName='"&oblog.filt_badstr(UserName)&"'")
		If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then
			Dim user_domainroot,Arr_domainroot,TEMP_domainroot
			TEMP_domainroot=Trim(oblog.CacheConfig(4))
			If InStr(TEMP_domainroot,"|")>0 Then
				Arr_domainroot=Split(TEMP_domainroot,"|")
				user_domainroot=Arr_domainroot(0)
			Else
				user_domainroot=TEMP_domainroot
			End If
			oblog.execute("update oblog_user set user_domain=userid where UserName='"&oblog.filt_badstr(UserName)&"'")
			oblog.execute("update oblog_user set user_domainroot='"&user_domainroot&"' where UserName='"&oblog.filt_badstr(UserName)&"'")
		End If
		oblog.CreateUserDir UserName,1
		rsreg.close
		set rsreg=Nothing
	Else
		ErrMsg=("用户名已存在，请更换重试！")
		FoundErr=True
		Exit Sub
	End If
End Sub
Function onlyEMail(mail)
onlyEMail=False
Dim rs, sql
	If Not IsObject(conn) Then link_database
	Set rs = Server.CreateObject("adodb.recordset")
	sql = "select * from [oblog_user] where useremail='" & Trim(mail) & "' "
	rs.Open sql, conn, 1, 1
	If rs.bof And rs.EOF Then onlyEMail=True
rs.Close: Set rs = Nothing
End Function
'oblog用户整合的登录函数
Sub ot_chklogin(UserName, PassWord, CookieDate)
	PassWord=XMLdom.documentElement.selectSingleNode("//password").text
	CookieDate=XMLdom.documentElement.selectSingleNode("//savecookie").text
	userip=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//userip").text)
	If UserName="" Then
		ErrMsg=ErrMsg&("用户名不能为空！")
		FoundErr=True
	End If
	If PassWord="" Then
		ErrMsg=ErrMsg&("密码不能为空！")
		FoundErr=True
	End If
	If FoundErr=True Then Exit Sub
	PassWord=md5(PassWord)
	Dim rs, sql
	If Not IsObject(conn) Then link_database
	Set rs = Server.CreateObject("adodb.recordset")
	sql = "select * from [oblog_user] where UserName='" & UserName & "' "
	rs.Open sql, conn, 1, 3
	If Not (rs.bof And rs.EOF) Then
			If rs("PassWord")=PassWord Then
				If rs("lockuser") = 1 Then
					rs.Close: Set rs = Nothing
					ErrMsg= ("对不起，你的ID已被锁定,不允许登录！"): FoundErr=True:Exit Sub
				Else
					rs("LastLoginIP") = userip
					rs("LastLoginTime") = oblog.ServerDate(Now())
					rs("LoginTimes") = rs("LoginTimes") + 1
					rs.Update
'					oblog.SaveCookie UserName, PassWord, CookieDate
'					SaveSession syskey,UserName,PassWord,""
					rs.Close: Set rs = Nothing
				End If
			Else
				rs.Close: Set rs = Nothing
				ErrMsg= ("用户名或者密码错误！"): FoundErr=True:Exit Sub
			End If
	Else
			rs.Close: Set rs = Nothing
			ErrMsg= ("用户名不存在！"): FoundErr=True:Exit Sub
	End If
End Sub
'oblog用户整合的登出函数
Sub LogoutUser()
	If FoundErr Then Exit Sub
	If cookies_domain <> "" Then
        Response.Cookies(cookies_name).domain = cookies_domain
    End If
	Response.Cookies(cookies_name).Path   =   blogdir
	Response.Cookies(cookies_name)("username")=oblog.CodeCookie("")
	Response.Cookies(cookies_name)("password")=oblog.CodeCookie("")
	Response.Cookies(cookies_name)("userurl")=oblog.CodeCookie("")
End Sub
'oblog用户整合的更新用户资料函数
Sub ModifyUserInfo()
	Call GetXML()
	If UserName="" Then
		ErrMsg=("用户名不能为空！")
		FoundErr=True
		Exit Sub
	End If
	Dim rs
	if not IsObject(conn) then Link_DataBase
	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select * from oblog_user where UserName='" & UserName & "'",conn,1,3
	If Not rs.eof Then
		If Email<>"" Then rs("useremail")=Email
		If PassWord<>"" Then rs("PassWord")=md5(PassWord)
		If Question<>"" Then  rs("Question")=Question
		If Answer<>""   Then  rs("Answer")=md5(Answer)
		If Sex<>"" And IsNumeric(Sex) Then  rs("Sex")=Sex
		If QQ<>"" And IsNumeric(QQ) Then  rs("QQ")=QQ
		If TrueName<>"" Then  rs("TrueName")=TrueName
		If Birthday<>"" Then  rs("Birthday")=Birthday
		If TelePhone<>"" And IsNumeric(TelePhone) Then  rs("tel")=TelePhone
		If HomePage<>"" Then  rs("HomePage")=HomePage
		If MSN<>"" Then  rs("MSN")=MSN
		If province<>"" Then  rs("province")=province
		If city<>"" Then  rs("city")=city
		If address<>"" Then  rs("address")=address
		If UserStatus<>"" Then
			If UserStatus=0 Then
				rs("Lockuser")=0
			Else
				rs("Lockuser")=1
			End If
		End If
		rs.update
		rs.close
'	Else
'		ErrMsg=("用户名不存在")
'		FoundErr=True
'		Exit Sub
	End If
	set rs=Nothing
End Sub
'oblog用户整合的删除用户函数
Sub DelUser()
	Dim rs,i
	If UserName="" Then
		ErrMsg= ("用户名不能为空(不能大于14小于4)！")
		FoundErr=True
		Exit Sub
	End If
	If InStr(UserName,",")>0 Then
		UserName=Split(UserName,",")
		For i=0 To UBound(UserName)
			deloneuser(UserName(i))
		Next
	Else
		deloneuser(UserName)
	End If
End Sub
'同上
Sub Deloneuser(UserName)
	If UserName="" Then
		ErrMsg=("用户名不能为空！")
		FoundErr=True
		Exit Sub
	End If
	Dim rs,fso,f,uname,udir,userid
	Set rs=oblog.execute("select user_dir,UserName,user_folder,userid from oblog_user where UserName='" & UserName & "'")
	If Not rs.eof Then
		udir=rs(0)
		uname=rs(1)
		userid=rs(3)
		Set fso=Server.CreateObject(oblog.CacheCompont(1))
		If fso.FolderExists(Server.MapPath(blogdir & udir&"/"&rs("user_folder"))) then
			Set f = fso.GetFolder(Server.MapPath(blogdir & udir&"/"&rs("user_folder")))
			f.delete True
		End If
		Set f=Nothing
		Set fso=Nothing
		oblog.execute("delete from oblog_log where userid="&userid)
		oblog.execute("delete from oblog_comment where userid="&userid)
		oblog.execute("delete from oblog_message where userid="&userid)
		oblog.execute("delete from oblog_subject where userid="&userid)
		oblog.execute("delete from oblog_user where userid=" & userid)
		oblog.execute("delete from oblog_upfile where userid=" & userid)
		oblog.execute("delete from oblog_friend where userid=" & userid)
		oblog.execute("update oblog_pm set dels=1 where sender='" &UserName&"'")
	End If
	Set rs=Nothing
End Sub
'oblog用户整合的获取用户信息函数
Sub getuserinfo()
	If UserName="" Then
		ErrMsg=("用户名不能为空！")
		FoundErr=True
		Exit Sub
	End If
	Dim rs,sql
	If Not IsObject(conn) Then link_database
	Set rs = Server.CreateObject("adodb.recordset")
	sql = "select * from [oblog_user] where UserName='" & UserName & "'"
	rs.Open sql, conn, 1, 1
	If Not rs.eof Then
			blogAPI.UserName=UserName
			blogAPI.PassWord=rs("password")
			blogAPI.CookieDate=CookieDate
			blogAPI.EMail=rs("useremail")
			blogAPI.Question=rs("question")
			blogAPI.Answer=rs("answer")
			blogAPI.Sex=rs("Sex")
			blogAPI.QQ=rs("QQ")
			blogAPI.MSN=rs("MSN")
			blogAPI.userstatus=rs("lockuser")
			blogAPI.truename=rs("TrueName")
			blogAPI.birthday=rs("Birthday")
			blogAPI.homepage=rs("HomePage")
			blogAPI.telephone=rs("Tel")
			blogAPI.address=rs("address")
			blogAPI.province=rs("province")
			blogAPI.city=rs("city")
			blogAPI.userip=oblog.userip
			blogAPI.GetUser
	Else
			ErrMsg=("用户名不存在！")
			FoundErr=True
			Exit Sub
	End If
	rs.close
	Set rs=Nothing
End Sub
'接收提交过来的XML数据
Sub GetXML()
	On Error Resume Next
	PassWord=XMLdom.documentElement.selectSingleNode("//password").text
	CookieDate=XMLdom.documentElement.selectSingleNode("//savecookie").text
	userip=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//userip").text)
	EMail=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//email").text)
'	Question=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//question").text)
'	Answer=XMLdom.documentElement.selectSingleNode("//answer").text
	Sex=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//gender").text)
	QQ=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//qq").text)
	MSN= oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//msn").text)
	userstatus=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//userstatus").text)
	truename=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//truename").text)
	birthday=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//birthday").text)
	homepage=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//homepage").text)
	telephone=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//telephone").text)
	province=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//province").text)
	city=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//city").text)
	address=oblog.filt_badstr(XMLdom.documentElement.selectSingleNode("//address").text)
End Sub
'验证提交信息的合法性，目前oblogMD5文件为16位，只能验证提交的位数再判定（考虑新版本增加参数）
Function ChkSyskey()
	ChkSyskey=True
	syskey=LCase(syskey)
	If Len(syskey)=32 Then
		If Mid(syskey,9,16)<>MD5(UserName&oblog_Key) Then
			ErrMsg=("安全码验证未通过！")
			FoundErr=True
			ChkSyskey=False
		End If
	ElseIf Len(syskey)=16 Then
		If syskey<>MD5(UserName&oblog_Key) Then
			ErrMsg=("安全码验证未通过！")
			FoundErr=True
			ChkSyskey=False
		End If
	Else
		ErrMsg=("安全码不合法！")
		FoundErr=True
		ChkSyskey=False
	End If
End Function
%>