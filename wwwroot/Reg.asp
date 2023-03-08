<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_email.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/inc_antispam.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="API/Class_API.asp"-->
<!--#include file="Inc/Rsa.Class.asp"-->

<%
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)

Call CheckBase()
'------------------------------------------------
Dim Action, sReg, sKeepKey,G_P_Show ,nTime,RegKey,ActionKey,sKeeyTime,SplRegkey
Action = Trim(Request("action"))

If Is_ot_User=1 Then
	If Not IsObject(conn) Then link_database
	Response.Redirect(ot_regurl)
	Set conn = Nothing
	Response.End()
End If
sKeepKey = Application(oblog.Cache_Name & "_RegKey")
If sKeepKey = "" Then
	Application(oblog.Cache_Name & "_RegKey") = GetDateCode(Now(),2) & RndPassword(12)
Else
	nTime = oblog.CacheConfig(60)
	If nTime < 30 Or nTime > 1440 Then nTime = 30
	If DateDiff("n", DeDateCode(Left(sKeepKey, 12)), Now) > nTime Then
		Application(oblog.Cache_Name & "_RegKey") = GetDateCode(Now(),2) & RndPassword(12)
	End If
End If
G_P_Show =  Replace (G_P_Show,"$show_title_list$", "新用户注册--"&oblog.cacheConfig(2) )

select Case action
	Case Application(oblog.Cache_Name & "_RegKey")
		Call Save
	Case "checkssn"
		Call checkssn
	Case "protocol"
		Call protocol
	Case Else
		Call ShowRegForm()

		'-====================================== 首先进行证书校验
'			If LCase(Request.ServerVariables("REQUEST_METHOD"))="post"  And Request.Cookies(cookies_name)("RegKey")<>"" Then
'			'ob_debug Request.Cookies(cookies_name)("RegKey"),1
'					SplRegkey=rsacode(Request.Cookies(cookies_name)("RegKey"),"de")
'					SplRegkey=Split(SplRegkey,"$",-1,1)
'					sKeeyTime= SplRegkey(0)
'					'5秒之内和600秒之外的数据直接丢弃。
'					If DateDiff("s", Date()&" "&sKeeyTime, Now) < 5 Or DateDiff("s", Date()&" "&sKeeyTime, Now) > 599 Then Call posterr(1)
'					If session(cookies_name&"RegKey")=SplRegkey(SplRegkey(1)) Then 	Call save
'					response.End
'			ElseIf LCase(Request.ServerVariables("REQUEST_METHOD"))="get" Then
'					Call ShowRegForm
'			Else
'					posterr(0)
'			End If
'-======================================
End select
G_P_Show=oblog.readfile("oblogstyle/reg/","reg.html")
G_P_Show = Replace(G_P_Show, "$show_title$",  "新用户注册－" &oblog.CacheConfig(2))
G_P_Show = Replace(G_P_Show, "$show_list$", sReg)
G_P_Show = Replace(G_P_Show,"$footer$", oblog.site_bottom)
Response.Write G_P_Show

'进行基础检测
Sub CheckBase()
    If oblog.CacheConfig(15) = 0 Then
       If oblog.CheckAdmin(0) = False Then
            oblog.adderrstr ("当前系统已关闭注册。")
            oblog.showerr
            Exit Sub
        End If
    End If
End Sub
	Sub posterr(e)
		Dim rearr,ajax
		set ajax=new AjaxXml
			If e=1 Then
			oblog.adderrstr ("您提交时间过长或者过短,我想您如果不是机器的话,那就刷新下重新注册一下吧.")
			Else
			oblog.adderrstr ("您也太神了，我想您不是太快了就是太慢了？|"&cookies_name)
			End If
			rearr=split(oblog.errstr&"$$1","$$")
			ajax.re(rearr)
			Response.end
	End Sub
'----------------------------------------------
Sub protocol()
	G_P_Show=oblog.readfile("oblogstyle/reg/","reg.html")
	G_P_Show = Replace(G_P_Show, "$show_list$", "当前位置：<a href='index.asp'>首页</a>→注册条款<hr />" & oblog.setup(9, 0))
	G_P_Show = Replace(G_P_Show, "$show_title$", oblog.CacheConfig(2) & "－注册条款")
	G_P_Show = Replace(G_P_Show,"$footer$", oblog.site_bottom)
	Response.Write G_P_Show
	Response.End
End Sub
Function MakeRsaStr()
	Dim RsaStr,Raddress,i
	Randomize
	Raddress=CStr(Int(5*Rnd))+2
	RsaStr=time()&"$"&Raddress+3
	For i=0 To 6
	RsaStr=RsaStr&"$"
	If i=Raddress Then RsaStr=RsaStr&"$"&Right(Application(oblog.Cache_Name & "_RegKey"),6)
	Next
	MakeRsaStr=RsaStr

End Function
Sub ShowRegForm()
'	'给客户端写入一个加密的标志
'	Dim RsaStr
'	RsaStr=MakeRsaStr()
'	'ob_debug rsacode(rsastr,"en"),1
'	session(cookies_name&"RegKey")=Right(Application(oblog.Cache_Name & "_RegKey"),6)
'	Response.Cookies(cookies_name)("RegKey")=Rsacode(RsaStr,"en")
'	Response.Cookies(cookies_name).Expires =DateAdd("n", 20, Now())
	Dim sUserType

	sUserType = "<select name=""usertype"" id=""usertype"" onBlur=""out_usertype()"" onChange=""out_usertype()"">"
	sUserType=sUserType&oblog.show_class("user",0,0)
	sUserType=sUserType&"</select>"
	sReg=sReg&"<form name=""oblogform"" method=""post"" action=""reg.asp?action="&Application(oblog.Cache_Name & "_RegKey")&""">" & vbcrlf
	sReg=sReg&"	<div id=""ob_reg"">" & vbcrlf
	sReg=sReg&"		<div class=""reg_content"">" & vbcrlf
	sReg=sReg&"			<fieldset>" &  vbcrlf
	sReg=sReg&"				<legend>基本信息</legend>" & vbcrlf
	If oblog.CacheConfig(17) = 1 Then
		sReg=sReg&"				<ul> " & vbcrlf
		sReg=sReg&"					<li class=""r_left""><label for=""obcode"">邀请码：</label></li>" & vbcrlf
		sReg=sReg&"					<li class=""r_right""><input name=""obcode"" id=""obcode"" type=""text"" size=""36"" maxlength=""32"" onFocus=""this.className='input_onFocus'"" onBlur=""this.className='input_onBlur'"" /></li>" & vbcrlf
		sReg=sReg&"				</ul>" & vbcrlf
	End If
	sReg=sReg&"				<ul>" & vbcrlf
	sReg=sReg&"					<li class=""r_left""><label for=""uname""><img src=""images/li_none.gif"" class=""okimg"" id=""d_uname_img"" />您的登录名：</label></li>" & vbcrlf
	sReg=sReg&"					<li class=""r_right""><input name=""username"" type=""text"" id=""uname"" size=""20"" maxlength=""30"" onFocus=""on_input('d_uname');this.className='input_onFocus'"" onBlur=""out_uname();this.className='input_onBlur'"" />" & vbcrlf
	sReg=sReg&"</li>" & vbcrlf
	sReg=sReg&"					<li class=""r_msg""><div id=""d_uname"" class=""d_default""></div></li>" & vbcrlf
	sReg=sReg&"				</ul>" & vbcrlf
	If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then
		sReg=sReg&"				<ul> " & vbcrlf
		sReg=sReg&"					<li class=""r_left""><label for=""domain""><img src=""images/li_none.gif"" class=""okimg"" id=""d_udomain_img"" />域名：</label></li>" & vbcrlf
		sReg=sReg&"					<li class=""r_right""><input name=""domain"" type=""text"" id=""domain"" size=""15"" maxlength=""30"" onFocus=""on_input('d_udomain');this.className='input_onFocus'"" onBlur=""out_udomain();this.className='input_onBlur'"" /> <select name=""user_domainroot"" id=""domainroot"">"&oblog.type_domainroot("",0) & "</select>" & vbcrlf
		sReg=sReg&"					</li>" & vbcrlf
		sReg=sReg&"					<li class=""r_msg""><div id=""d_udomain"" class=""d_default""></div></li>" & vbcrlf
	Else
		sReg=sReg&"<input name=""domain"" type=""hidden"" id=""domain"" size=""15"" maxlength=""30"" /><input type=""hidden"" name=""user_domainroot"" />" & vbcrlf
	End If
	sReg=sReg&"				</ul>" & vbcrlf
	sReg=sReg&"				<ul>" & vbcrlf
	sReg=sReg&"					<li class=""r_left""><span id=""chkssn_stat""></span></li>" & vbcrlf
	If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then
		sReg=sReg&"					<li class=""r_right""><input type=""button"" id=""ssnbotton"" onclick=""checkssn();"" value=""查看用户名、域名是否可用"" /></li>" & vbcrlf
	Else
		sReg=sReg&"					<li class=""r_right""><input type=""button"" id=""ssnbotton"" onclick=""checkssn();"" value=""查看用户名是否可用"" /></li>" & vbcrlf
	End If
	sReg=sReg&"				</ul>" & vbcrlf
	sReg=sReg&"			</fieldset>" & vbcrlf
	sReg=sReg&"		</div>" & vbcrlf
	sReg=sReg&"		<div class='reg_content'>" & vbcrlf
	sReg=sReg&"			<fieldset>" & vbcrlf
	sReg=sReg&"				<legend>安全资料</legend>" & vbcrlf
	sReg=sReg&"					<ul>" & vbcrlf
	sReg=sReg&"						<li class=""r_left""><label for=""upwd""><img src=""images/li_none.gif"" class=""okimg"" id=""d_upwd1_img"" />输入登录密码：</label></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right""><input name=""password"" type=""password"" id=""upwd"" size=""20"" maxlength=""12"" onKeyUp=""EvalPwdStrength(this.value);"" onFocus=""on_input('d_upwd1');this.className='input_onFocus'"" onBlur=""out_upwd1();this.className='input_onBlur'"" /> "
	sReg=sReg&"						<li class=""r_msg""><div id=""d_upwd1"" class=""d_default""></div></li>" & vbcrlf
	sReg=sReg&"					</ul>" & vbcrlf
	sReg=sReg&"					<ul>" & vbcrlf
	sReg=sReg&"						<li class=""r_left""></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right"">" & vbcrlf
	sReg=sReg&"							<div id=""pws"" class=""ob_pws"">" & vbcrlf
	sReg=sReg&"								<div id=""idSM1"" class=""ob_pws0""><span style=""font-size:1px"">&nbsp;</span><span id=""idSMT1"">弱</span></div>" & vbcrlf
	sReg=sReg&"								<div id=""idSM2"" class=""ob_pws0""  style=""border-left:solid 1px #DEDEDE""><span style=""font-size:1px"">&nbsp;</span><span id=""idSMT2"">中</span></div>" & vbcrlf
	sReg=sReg&"								<div id=""idSM3"" class=""ob_pws0"" style=""border-left:solid 1px #DEDEDE""><span style=""font-size:1px"">&nbsp;</span><span id=""idSMT3"">强</span></div>" & vbcrlf
	sReg=sReg&"							</div>" & vbcrlf
	sReg=sReg&"						</li>" & vbcrlf
	sReg=sReg&"					</ul>" & vbcrlf
	sReg=sReg&"					<ul> " & vbcrlf
	sReg=sReg&"						<li class=""r_left""><label for=""repassword""><img src='images/li_none.gif' class='okimg' id='d_upwd2_img' />登录密码确认：</label></li>" & vbcrlf
	sReg=sReg&"						<li class='r_right'><input name=""repassword"" type=""password"" id=""repassword"" size=""20"" maxlength=""20"" onFocus=""on_input('d_upwd2');this.className='input_onFocus'"" onBlur=""out_upwd2();this.className='input_onBlur'"" /></li>" & vbcrlf
	sReg=sReg&"						<li class='r_msg'><div id='d_upwd2' class='d_default'></div></li></ul>" & vbcrlf
	sReg=sReg&"					</ul>" & vbcrlf
	sReg=sReg&"					<ul> " & vbcrlf
	sReg=sReg&"						<li class=""r_left""><label for=""question""><img src=""images/li_none.gif"" class=""okimg"" id=""d_question_img"" />密码提示问题：</label></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right"">" & vbcrlf
	sReg=sReg&"							<select id=""question"" name=""question"" onBlur=""out_question()"" onChange=""out_question()"">" & vbcrlf
	sReg=sReg&"								<option selected value="""">--请您选择--</option>" & vbcrlf
	sReg=sReg&"								<option value=""我的宠物名字？"">我的宠物名字？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最好的朋友是谁？"">我最好的朋友是谁？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜爱的颜色？"">我最喜爱的颜色？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜爱的电影？"">我最喜爱的电影？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜爱的影星？"">我最喜爱的影星？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜爱的歌曲？"">我最喜爱的歌曲？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜爱的食物？"">我最喜爱的食物？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最大的爱好？"">我最大的爱好？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我中学校名全称是什么？"">我中学校名全称是什么？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我的座右铭是？"">我的座右铭是？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜欢的小说的名字？"">我最喜欢的小说的名字？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜欢的卡通人物名字？"">我最喜欢的卡通人物名字？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我母亲/父亲的生日？"">我母亲/父亲的生日？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最欣赏的一位名人的名字？"">我最欣赏的一位名人的名字？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜欢的运动队全称？"">我最喜欢的运动队全称？</option>" & vbcrlf
	sReg=sReg&"								<option value=""我最喜欢的一句影视台词？"">我最喜欢的一句影视台词？</option>" & vbcrlf
	sReg=sReg&"							</select>" & vbcrlf
	sReg=sReg&"						</li>" & vbcrlf
	sReg=sReg&"						<li class=""r_msg""><div id=""d_question"" class=""d_default""></div></li>" & vbcrlf
	sReg=sReg&"					</ul>" & vbcrlf
	sReg=sReg&"					<ul> " & vbcrlf
	sReg=sReg&"						<li class=""r_left""><label for=""answer""><img src=""images/li_none.gif"" class=""okimg"" id=""d_an_img"" />密码提示答案：</label></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right""><input name=""answer"" type=""text"" id=""answer"" size=""30"" maxlength=""30"" onFocus=""on_input('d_an');this.className='input_onFocus'"" onBlur=""out_an();this.className='input_onBlur'"" /></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_msg""><div id=""d_an"" class=""d_default""></div></li>" & vbcrlf
	sReg=sReg&"					</ul>" & vbcrlf
	sReg=sReg&"				</fieldset>" & vbcrlf
	sReg=sReg&"			</div>" & vbcrlf
	sReg=sReg&"			<div class=""reg_content"">" & vbcrlf
	sReg=sReg&"				<fieldset>" & vbcrlf
	sReg=sReg&"					<legend>个人资料</legend>" & vbcrlf
	sReg=sReg&"						<ul>" & vbcrlf
	sReg=sReg&"						<li class=""r_left""><label for=""email""><img src=""images/li_none.gif"" class=""okimg"" id=""d_email_img"" />电子邮箱：</label></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right""><input name=""email"" id=""email"" type=""text"" size=""30"" maxlength=""32"" onFocus=""on_input('d_email');this.className='input_onFocus'"" onBlur=""out_email();this.className='input_onBlur'"" /></li>"
	sReg=sReg&"						<li class=""r_msg""><div id=""d_email"" class=""d_email""></div></li>" & vbcrlf
	sReg=sReg&"					</ul>" & vbcrlf
	sReg=sReg&"					<ul> " & vbcrlf
	sReg=sReg&"						<li class=""r_left""><label for=""city""><img src=""images/li_none.gif"" class=""okimg"" id=""d_city_img"" />地区(省/市)：</label></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right"">"&show_city()  & "</li>" & vbcrlf
	sReg=sReg&"						<li class=""r_msg""><div id=""d_city"" class=""d_city""></div></li>" & vbcrlf
	sReg=sReg&"					</ul> " & vbcrlf
	sReg=sReg&"					<ul> " & vbcrlf
	sReg=sReg&"						<li class=""r_left""><label for=""blogname""><img src=""images/li_none.gif"" class=""okimg"" id=""d_blogname_img"" />Blog名称：</label></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right""><input name=""blogname"" id=""blogname"" type=""text"" size=""30"" maxlength=""30"" onFocus=""on_input('d_blogname');this.className='input_onFocus'"" onBlur=""out_blogname();this.className='input_onBlur'"" /></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_msg""><div id=""d_blogname"" class=""d_default""></div></li>" & vbcrlf
	sReg=sReg&"					</ul> " & vbcrlf
	sReg=sReg&"					<ul> " & vbcrlf
	sReg=sReg&"						<li class=""r_left""><label for=""UserType""><img src=""images/li_none.gif"" class=""okimg"" id=""d_usertype_img"" />Blog类别：</label></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right"">"&sUserType&"</li>" & vbcrlf
	sReg=sReg&"						<li class=""r_msg""><div id=""d_usertype"" class=""d_default""></div></li>" & vbcrlf
	sReg=sReg&"					</ul> " & vbcrlf
	sReg=sReg&"					<ul> " & vbcrlf
	sReg=sReg&"						<li class=""r_left""><img src=""images/li_none.gif"" class=""okimg"" id=""d_passregtext_img"" />注册条款：</li>" & vbcrlf
	sReg=sReg&"						<li class=""r_right""><label><input name=""passregtext"" id=""passregtext"" type=""radio"" value=""1"" checked onClick=""out_passregtext();"" />同意</label>　<label><input type=""radio"" name=""passregtext"" id=""passregtext1"" value=""0"" onClick=""out_passregtext();"" />不同意</label>　<a href=""#showpassregtext"" onClick=""return doMenu('showpassregtext');"">查看注册条款</a></li>" & vbcrlf
	sReg=sReg&"						<li class=""r_msg""><div id=""d_passregtext"" class=""d_default""></div></li>" & vbcrlf
	sReg=sReg&"					</ul>" & vbcrlf
	sReg=sReg&"			<div id=""showpassregtext"" name=""showpassregtext"" style=""display: none; "">" & vbcrlf
	sReg=sReg& oblog.setup(9, 0) & vbcrlf
	sReg=sReg&"			</div>" & vbcrlf
	If oblog.CacheConfig(16) = 1 Then
		sReg=sReg&"					<ul> " & vbcrlf
		sReg=sReg&"						<li class=""r_left""><label for=""codestr"">验证码：</label></li>" & vbcrlf
		sReg=sReg&"						<li class=""r_right""><input name=""codestr"" id=""codestr"" type=""text"" onFocus=""this.className='input_onFocus'"" onBlur=""this.className='input_onBlur'"" size=""4"" maxlength=""20""> "&oblog.getcode&"</li>" & vbcrlf
		sReg=sReg&"					</ul>" & vbcrlf
	End If
	sReg=sReg&"				</fieldset>" & vbcrlf
	sReg=sReg&"			</div>" & vbcrlf
	sReg=sReg&"		<ul>" & vbcrlf
	sReg=sReg&"			<li class=""r_left""></li>" & vbcrlf
	sReg=sReg&"			<li><input type=""button"" name=""submit"" value=""OK!确认提交注册信息"" onclick=""chk_reg();"" id=""regbotton""><span id=""save_stat""></span></li>" & vbcrlf
	sReg=sReg&"		</ul>" & vbcrlf
	sReg=sReg&"	</div>" & vbcrlf
	sReg=sReg&"</form>" & vbcrlf
End Sub

Sub Save()
	If oblog.ChkPost() = False Then
        oblog.adderrstr ("系统不允许从外部提交！")
    End If
    Dim rsreg, sql, ajax, buttonface, rearr
    Dim regusername, regpassword, sex, question, answer, email, reguserlevel, userispass, blogname, usertype, nickname
    Dim re_regpassword, user_domain, user_domainroot
	Dim chk_regname
	chk_regname=oblog.chk_regname(regusername)
	buttonface=2
    set ajax=new AjaxXml
    chk_regtime()
    If oblog.CacheConfig(16)=1 Then
        If Not oblog.codepass Then
			oblog.adderrstr ("验证码错误，请刷新后重新输入！")
			rearr=split(oblog.errstr&"$$1","$$")
			ajax.re(rearr)
			Response.end
		end if
    End If
    If oblog.CacheConfig(17)=1 Then
    	if oblog.CheckOBCode(Request("obcode"),0)=false Then
    		oblog.adderrstr ("邀请码错误或已经被使用！")
    		rearr=split(oblog.errstr&"$$1","$$")
				ajax.re(rearr)
				Response.end
			End If
		End If
    regusername = oblog.filt_badstr(Trim(Request("username")))
    regpassword = Trim(Request("password"))
    re_regpassword = Trim(Request("repassword"))
    email = Trim(Request("email"))
    question = Trim(Request("question"))
    answer = Trim(Request("answer"))
    blogname = Trim(Request("blogname"))
    usertype = CLng(Request("usertype"))
    user_domain = LCase(Trim(Request("domain")))
    user_domainroot = Trim(Request("user_domainroot"))
    If regusername = "" Or oblog.strLength(regusername) > 14 Or oblog.strLength(regusername) < 4 Then oblog.adderrstr ("用户名不能为空(不能大于14小于4)！")
	if chk_regname>0 then
'		if chk_regname = 1 Then oblog.adderrstr("用户名不合规范，只能使用小写字母，数字及下划线！")
		if chk_regname = 2 Then oblog.adderrstr("用户名中含有系统不允许的字符！")
		if chk_regname = 3 Then oblog.adderrstr("用户名中含有系统保留注册的字符！")
		if chk_regname = 4 Then oblog.adderrstr("用户名中不允许全部为数字！")
	End If
	If oblog.CacheConfig(6) <> "1" Then
		If oblog.chkdomain(regusername) = False Then oblog.adderrstr  ("用户名不合规范，只能使用小写字母，数字及下划线！")
	End If
	If oblog.CacheConfig(4) <>"" And oblog.CacheConfig(5) = 1 Then
        If user_domain = "" Or oblog.strLength(user_domain) > 14 Then oblog.adderrstr  ("域名不能为空(不能大于14个字符)！")
        If user_domain <> Request("old_userdomain") And oblog.strLength(user_domain) < 4 Then oblog.adderrstr  ("域名不能小于4个字符！")
        If oblog.chk_regname(user_domain) Then oblog.adderrstr  ("此域名系统不允许注册！")
        If oblog.chk_badword(user_domain) > 0 Then oblog.adderrstr  ("域名中含有系统不允许的字符！")
        If oblog.chkdomain(user_domain) = False Then oblog.adderrstr  ("域名不合规范，只能使用小写字母，数字！")
        If user_domainroot = "" Then oblog.adderrstr  ("域名根不能为空！")
		If oblog.CheckDomainRoot(user_domainroot,0) = False Then oblog.adderrstr  ("域名根不合法！")
    End If
    If regpassword = "" Or oblog.strLength(regpassword) > 14 Or oblog.strLength(regpassword) < 4 Then oblog.adderrstr  ("密码不能为空(不能大于14小于4)！")
    If re_regpassword = "" Then oblog.adderrstr  ("重复密码不能为空！")
    If regpassword <> re_regpassword Then oblog.adderrstr  ("两次输入密码不同！")
    If question = "" Or oblog.strLength(question) > 50 Then oblog.adderrstr  ("找回密码提示问题不能为空(不能大于50)！")
    If answer = "" Or oblog.strLength(answer) > 50 Then oblog.adderrstr  ("找回密码问题答案不能为空(不能大于50)！")
    If blogname = "" Or oblog.strLength(blogname) > 50 Then oblog.adderrstr  ("blog名不能为空(不能大于50字符)！")
    If oblog.chk_badword(blogname) > 0 Then oblog.adderrstr  ("blog名中含有系统不允许的字符！")
    If InStr(regusername, "=") > 0 Or InStr(regusername, "%") > 0 Or InStr(regusername, Chr(32)) > 0 Or InStr(regusername, "?") > 0 Or InStr(regusername, "&") > 0 Or InStr(regusername, ";") > 0 Or InStr(regusername, ",") > 0 Or InStr(regusername, "'") > 0 Or InStr(regusername, ",") > 0 Or InStr(regusername, Chr(34)) > 0 Or InStr(regusername, Chr(9)) > 0 Or InStr(regusername, "") > 0 Or InStr(regusername, "$") > 0 Or InStr(regusername, ".") > 0  Then oblog.adderrstr  ("用户名中含有非法字符！")
    '进行重复性判断22/47/25
    If oblog.CacheConfig(22)="1" Then
    	Set rsreg=oblog.execute("select Count(userid) From oblog_user Where useremail='" & ProtectSQL(email) & "'")
    	If rsreg(0)>0 Then
    		oblog.adderrstr  ("您使用的Email: " & email & " 已被他人使用，请更换其他Email")
    	End If
    	rsreg.Close
	End If
	If oblog.CacheConfig(48)="1" Then
		Set rsreg=oblog.execute("select Count(userid) From oblog_user Where blogname='" & ProtectSQL(blogname) & "'")
    	If rsreg(0)>0 Then
    		oblog.adderrstr  ("您使用的博客名称: " & blogname & " 已被他人使用，请更换博客名称")
    	End If
    	rsreg.Close
	End If
    '进行IP控制
	Dim sIP
    sIP=oblog.userip
    If oblog.CacheConfig(21)>"0" And oblog.ChkWhiteIP(sIP) = False Then
		sql="select Count(userid) from oblog_user where regip='"& sIP &"' And "
		If Is_Sqldata = 0 Then
			sql = sql & " Datediff('n',adddate,Now())<=60"
		Else
			sql = sql & " adddate BETWEEN DATEADD(Minute,-60,GETDATE()) AND GETDATE()"
		End if
		Set rsreg = oblog.execute(sql)
		If rsreg(0) > Int(oblog.CacheConfig(21)) Then
			oblog.KillIP(sIP)
			oblog.adderrstr  ("您的IP因为恶意注册被临时禁止")
			rsreg.Close
			rearr=split(Replace(oblog.errstr,"_","<br />")&"$$2$$index","$$")
			ajax.re(rearr)
			Response.end
		End If
		rsreg.Close
	End If
    If oblog.CacheConfig(14) > "0" And oblog.ChkWhiteIP(sIP) = False Then
		sql="select Count(userid) from oblog_user where regip='"& sIP &"' And "
		If Is_Sqldata = 0 Then
			sql = sql & " Datediff('h',adddate,Now())<=24"
		Else
			sql = sql & " adddate BETWEEN DATEADD(Hour,-24,GETDATE()) AND GETDATE()"
		End IF
		Set rsreg = oblog.execute(sql)
		If rsreg(0) > Int(oblog.CacheConfig(14)) Then
		   '进行IP屏蔽
			oblog.KillIP(sIP)
			'进行批量屏蔽
			sql=""
		If Is_Sqldata = 0 Then
			sql = sql & " Datediff('h',adddate,Now())<=24"
		Else
			sql = sql & " adddate BETWEEN DATEADD(Hour,-24,GETDATE()) AND GETDATE()"
		End IF
			oblog.execute("Update [oblog_user] Set user_level=6 Where regip='"&oblog.userip&"' and "&sql)
			oblog.adderrstr  ("您的IP因为恶意注册而被系统禁止")
			rsreg.Close
			rearr=split(Replace(oblog.errstr,"_","<br />")&"$$2$$index","$$")
			ajax.re(rearr)
			Response.end
		End If
		rsreg.Close
	End If
    If user_domain <> "" Then
        Set rsreg = oblog.execute("select userid from oblog_user where user_domain='" & oblog.filt_badstr(user_domain) & "' and user_domainroot='" & oblog.filt_badstr(user_domainroot) & "'")
        If Not rsreg.EOF Or Not rsreg.bof Then oblog.adderrstr  ("系统中已经有这个域名存在，请更改域名！")
    End If
    If oblog.errstr <> "" Then
		rearr=split(Replace(oblog.errstr,"_","<br />")&"$$1","$$")
		ajax.re(rearr)
		Response.end
	end if
    '是否需要审核
    If oblog.CacheConfig(18) = 1 Then reguserlevel = 6 Else reguserlevel = 7

	If API_Enable Then
		Dim blogAPI
		Set blogAPI = New DPO_API_OBLOG
		blogAPI.LoadXmlFile True
		blogAPI.UserName=regusername
		blogAPI.PassWord=regpassword
		blogAPI.EMail=email
		blogAPI.Question=Question
		blogAPI.Answer=Answer
		blogAPI.userip=oblog.userip
		blogAPI.UserStatus=0
		Call blogAPI.ProcessMultiPing("reguser")
		Set blogAPI=Nothing
		Dim strUrl,i,turl
		For i=0 To UBound(aUrls)
			strUrl=Lcase(aUrls(i))
			If Left(strUrl,7)="http://" Then
				turl=strUrl&"?syskey="&MD5(regusername&oblog_Key)&"&username="&regusername&"&password="&MD5(regpassword)&"&savecookie=1@@@"& turl
			End If
		Next
		session("turl")=turl
	End If

	Dim TruePassWord,IsEmailReg
	IsEmailReg=False
	TruePassWord = RndPassword(16)
    If Not IsObject(conn) Then link_database
    Set rsreg = Server.CreateObject("adodb.recordset")
    rsreg.open "select * from [oblog_user] where username='" & regusername & "'", conn, 1, 3
    If rsreg.EOF Then
    	rsreg.addnew
        rsreg("username") = regusername
        rsreg("password") = MD5(regpassword)
		rsreg("TruePassWord") = TruePassWord
        If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then
            rsreg("user_domain") = user_domain
            rsreg("user_domainroot") = user_domainroot
        End If
        rsreg("question") = question
        rsreg("answer") = MD5(answer)
        rsreg("useremail") = email

		If oblog.CacheConfig(88) = "0" Then
			 rsreg("user_level") = reguserlevel
		Else
			rsreg("user_level") = 6
			rsreg("is_log_default_hidden") = 1
			IsEmailReg=True
		End If
		If oblog.CacheConfig(89) = 1 Or reguserlevel = 6  Then rsreg("is_log_default_hidden") = 1
        rsreg("user_isbest") = 0
        rsreg("blogname") = blogname
        rsreg("user_classid") = usertype
        'rsreg("nickname")=nickname
        rsreg("province") = Request("province")
        rsreg("city") = Request("city")
        rsreg("adddate") = oblog.ServerDate(Now())
        rsreg("regip") = oblog.userip
        rsreg("lastloginip") = oblog.userip
        rsreg("lastlogintime") = oblog.ServerDate(Now())
        rsreg("user_dir") =oblog.setup(8,0)
        rsreg("user_folder") = regusername
        rsreg("user_group") = oblog.defaultGroup
        rsreg("scores") = oblog.cacheScores(1)
        rsreg("newbie") = 1
		rsreg("isdigg") = 1
		if oblog.CacheConfig(40)=1 then rsreg("comment_isasc")=1
        rsreg.Update
		Session("chk_regtime") = Now()
        oblog.execute ("update oblog_setup set user_count=user_count+1")
        oblog.execute ("update oblog_groups set g_members=g_members+1 WHERE groupid = " &oblog.defaultGroup)
        If oblog.CacheConfig(58) = "0" Or oblog.CacheConfig(6) = "1" Then
            oblog.execute ("update oblog_user set user_folder=userid where username='" & regusername & "'")
        End If
        If oblog.CacheConfig(17) = "1" Then
			Dim tid1,tid2
        	'获取Userid
        	Set rsreg=oblog.Execute("select userid From oblog_user where username='" & regusername & "'")
			If Not rsreg.Eof Then
				tid1=rsreg(0)
        		oblog.Execute("Update oblog_obcodes Set istate=1,useip='" &oblog.userip & "',usetime='" & Now & "',useuser=" & tid1 & " Where obcode='" &oblog.filt_badstr(Request("obcode")) & "'"  )
			End if
        	rsreg.Close
			'增加积分
			Set rsreg = oblog.execute ("select creatuser FROM oblog_obcodes Where obcode='" &oblog.filt_badstr(Request("obcode")) & "'"  )
			If Not rsreg.Eof Then
				tid2=rsreg(0)
				oblog.GiveScore "",oblog.cacheScores(2),tid2
			End if
			rsreg.Close
			'互相加好友
			oblog.execute("insert into [oblog_friend] (userid,friendid,isblack) values ("&tid1&","&tid2&",0)")
			oblog.execute("insert into [oblog_friend] (userid,friendid,isblack) values ("&tid2&","&tid1&",0)")
        End If
        If oblog.CacheConfig(59) = "1" Then
			oblog.CreateUserDir regusername, 1
			If oblog.CacheConfig(17) = "1" Then
				dim blog
				set blog=new class_blog
				blog.userid=tid1
				blog.update_friends tid1
				blog.userid=tid2
				blog.update_friends tid2
				set blog=nothing
			End if
			'自动选择默认用户模板
			Set rsreg=oblog.Execute("select userid From oblog_user where username='" & regusername & "'")
			C_Template rsreg(0)
			rsreg.Close
		End if
        If oblog.CacheConfig(18) = 1 or IsEmailReg Then
            rearr=rearr&"注册成功，但当前系统设置为需要通过审核，您暂时没有管理权限！请接收您的激活信件。<br />"
			rearr=rearr&"$$2$$index"
        Else
            oblog.savecookie regusername,TruePassWord,0
            rearr=rearr&"恭喜！您已经注册成功！<br />"
            rearr=rearr&"现在将转到管理后台让您选择喜欢的页面风格。</a><br />"
			rearr=rearr&"$$2$$user_index"

			If API_Enable Then
				rearr=rearr&"$$"&MD5(regusername & oblog_Key )&"$$"&regusername&"$$ "&MD5(regpassword)
			End If

        End If


		ajax.re(split(rearr,"$$"))
				'''''''''''''''''''''''''''''
		If IsEmailReg or reguserlevel =6  Then
		Dim ma

			Set ma=new Oblog_email
		Call ma.SendValidAccountMail(regusername,email)
		End If

		''''''''''''''''''''''''''''
		Response.end
    Else
        oblog.adderrstr ("系统中已经有这个用户名存在，请更改用户名！")
        ajax.re(split(Replace(oblog.errstr,"_","<br />")&"$$1","$$"))
		Response.end
        Exit Sub
    End If
    rsreg.Close
    Set rsreg = Nothing
End Sub
Sub chk_regtime()
    Dim lasttime,rearr,ajax
	set ajax=new AjaxXml
    lasttime = Session("chk_regtime")
    If IsDate(lasttime) Then
        If DateDiff("s", lasttime, Now()) < CLng(oblog.CacheConfig(20)) Then
			oblog.adderrstr (oblog.CacheConfig(20) & "秒后才能重复注册。")
			rearr=split(oblog.errstr&"$$1","$$")
			ajax.re(rearr)
			Response.End
        End If
    End If
End Sub

Sub checkssn()
	Dim ajax,rearr,msgstr,buttomface
	dim regusername,user_domain,user_domainroot,email
	Dim chk_regname
	buttomface=2
	regusername=oblog.filt_badstr(Trim(Request("username")))
	user_domain=oblog.filt_badstr(Trim(Request("domain")))
	user_domainroot=oblog.filt_badstr(Trim(Request("domainroot")))
	email=oblog.filt_badstr(Trim(Request("email")))
	chk_regname=oblog.chk_regname(regusername)
	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("用户名不能为空(不能大于14小于4)！")
	if chk_regname>0 then
'		if chk_regname = 1 Then oblog.adderrstr("用户名不合规范，只能使用小写字母，数字及下划线！")
		if chk_regname = 2 Then oblog.adderrstr("用户名中含有系统不允许的字符！")
		if chk_regname = 3 Then oblog.adderrstr("用户名中含有系统保留注册的字符！")
		if chk_regname = 4 Then oblog.adderrstr("用户名中不允许全部为数字！")
	End If
	If oblog.CacheConfig(6) <> "1" Then
		If oblog.chkdomain(regusername) = False Then oblog.adderrstr  ("用户名不合规范，只能使用小写字母，数字及下划线！")
	End if
	if oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) then
		if user_domain="" or oblog.strLength(user_domain)>20  then oblog.adderrstr("域名不能为空(不能大于14个字符)！")
		if user_domain<>Request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("域名不能小于4个字符！")
		if oblog.chk_regname(user_domain) then oblog.adderrstr("此域名系统不允许注册！")
		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("域名中含有系统不允许的字符！")
		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("域名不合规范，只能使用小写字母，数字及下划线！")
		if user_domainroot="" then oblog.adderrstr("域名根不能为空！")
	end If
	If API_Enable Then
		Dim blogAPI
		Set blogAPI = New DPO_API_OBLOG
		blogAPI.LoadXmlFile True
		blogAPI.UserName=regusername
		blogAPI.email=email
		Call blogAPI.ProcessMultiPing("checkname")
	End If
	If oblog.errstr<>"" Then
		msgstr=Replace(oblog.errstr,"_","<br />")
		buttomface=1
	else
		dim rs
		set rs=oblog.execute("select userid from oblog_user where username='"&regusername&"'")
		if not rs.eof then
			msgstr="对不起，<strong>"&regusername&"</strong>此用户名已存在,请更换！<br />"
			buttomface=1
		Else
			If API_Enable Then
				If blogAPI.FoundErr=False Then
					msgstr="恭喜，<strong>"&regusername&"</strong>此用户名可使用！<br />"
				End If
			Else
				msgstr="恭喜，<strong>"&regusername&"</strong>此用户名可使用！<br />"
			End If
		end if

		if oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) Then
			set rs=oblog.execute("select userid from oblog_user where user_domain='"&user_domain&"' and user_domainroot='"&user_domainroot&"'")
			if not rs.eof then
				msgstr=msgstr&"对不起，<strong>"&user_domain&"."&user_domainroot&"</strong>此域名已存在,请更换！<br />"
				buttomface=1
			else
				msgstr=msgstr&"恭喜，<strong>"&user_domain&"."&user_domainroot&"</strong>此域名可使用！<br />"
			end If
			If oblog.CheckDomainRoot(user_domainroot,0) = False Then msgstr=("域名根不合法！"):buttomface=1
		end if
	End If
	set rs=nothing
	If API_Enable Then Set blogAPI=Nothing
	rearr=split(msgstr&"$$"&buttomface,"$$")
	set ajax=new AjaxXml
	ajax.re(rearr)
	Response.End()
End Sub

Sub C_Template(userid)
	Dim rs,rsskin
	Set rsskin=oblog.execute("select skinmain,skinshowlog,id from oblog_userskin where isdefault=1")
	If rsskin.EOF Then
		Set rsskin=oblog.execute("select top 1 skinmain,skinshowlog,id from oblog_userskin order by id desc")
	End if
	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select user_skin_main,user_skin_showlog,defaultskin from [oblog_user] where userid="&userid,conn,1,3
	rs(0) = rsskin(0)
	rs(1) = rsskin(1)
	rs(2) = rsskin(2)
	Set rsskin=Nothing
	rs.update
	rs.close
	Set rs=Nothing
	Dim blog
	Set blog=new class_blog
	blog.userid = userid
	blog.update_index 0
	blog.update_message 0
	blog.CreateFunctionPage
	Set blog=Nothing
	oblog.execute "Update oblog_user Set newbie=0 Where userid=" & userid
End Sub


function show_city()
		Dim tmpstr
        tmpstr = "<select onchange='setcity();out_city();' name='province' >"
        tmpstr = tmpstr & "<option value=''>--请选择省份--</option>"
        tmpstr = tmpstr & "<option "
        tmpstr = tmpstr & "value=安徽>安徽</option> <option value=北京>北京</option> "
        tmpstr = tmpstr & "<option value=重庆>重庆</option> <option "
        tmpstr = tmpstr & "value=福建>福建</option> <option value=甘肃>甘肃</option> "
        tmpstr = tmpstr & "<option value=广东>广东</option> <option "
        tmpstr = tmpstr & "value=广西>广西</option> <option value=贵州>贵州</option> "
        tmpstr = tmpstr & "<option value=海南>海南</option> <option "
        tmpstr = tmpstr & "value=河北>河北</option> <option value=黑龙江>黑龙江</option> "
        tmpstr = tmpstr & "<option value=河南>河南</option> <option "
        tmpstr = tmpstr & "value=香港>香港</option> <option value=湖北>湖北</option> "
        tmpstr = tmpstr & "<option value=湖南>湖南</option> <option "
        tmpstr = tmpstr & "value=江苏>江苏</option> <option value=江西>江西</option> "
        tmpstr = tmpstr & "<option value=吉林>吉林</option> <option "
        tmpstr = tmpstr & "value=辽宁>辽宁</option> <option value=澳门>澳门</option>"
        tmpstr = tmpstr & "<option value=内蒙古>内蒙古</option> <option "
        tmpstr = tmpstr & "value=宁夏>宁夏</option> <option value=青海>青海</option> "
        tmpstr = tmpstr & "<option value=山东>山东</option> <option "
        tmpstr = tmpstr & "value=上海>上海</option> <option value=山西>山西</option> "
        tmpstr = tmpstr & "<option value=陕西>陕西</option> <option "
        tmpstr = tmpstr & "value=四川>四川</option> <option value=台湾>台湾</option> "
        tmpstr = tmpstr & "<option value=天津>天津</option> <option "
        tmpstr = tmpstr & "value=新疆>新疆</option> <option value=西藏>西藏</option> "
        tmpstr = tmpstr & "<option value=云南>云南</option> <option "
        tmpstr = tmpstr & "value=浙江>浙江</option> <option "
        tmpstr = tmpstr & "value=海外>海外</option></select>"
        tmpstr = tmpstr & " <select name='city' id = 'city' >"
        tmpstr = tmpstr & "</select>"
        tmpstr = tmpstr & "<script src=""inc/getcity.js""></script>"
        tmpstr = tmpstr & "<script>initprovcity('','');</script>"
        show_city = tmpstr
End Function

%>
<script src="inc/main.js"></script>
<script language="javascript">
<!--
var msg	;
var li_ok='images/li_ok.gif';
var li_err='images/li_err.gif'
var bname_m=false;
function init_reg(){
	msg=new Array(
	"请输入4-14位字符，英文、数字的组合。",
	"请输入4-14位字符，英文、数字的组合。",
	"请输入6位以上字符，不允许空格。",
	"请重复输入上面的密码。",
	"请选择密码提示问题。",
	"5个字符、数字或3个汉字以上（包括6个）。",
	"请输入您常用的电子邮箱地址。",
	"请输入您的blog名称。",
	"请选择您所在的地区。",
	"请选择您的blog类别。",
	"只有同意注册条款才能完成注册。"
	)
	document.getElementById("d_uname").innerHTML=msg[0];
	<%If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then%>
	document.getElementById("d_udomain").innerHTML=msg[1];
	<%end if%>
	document.getElementById("d_upwd1").innerHTML=msg[2];
	document.getElementById("d_upwd2").innerHTML=msg[3];
	document.getElementById("d_question").innerHTML=msg[4];
	document.getElementById("d_an").innerHTML=msg[5];
	document.getElementById("d_email").innerHTML=msg[6];
	document.getElementById("d_blogname").innerHTML=msg[7];
	document.getElementById("d_city").innerHTML=msg[8];
	document.getElementById("d_usertype").innerHTML=msg[9];
}
init_reg();
function on_input(objname){
	var strtxt;
	var obj=document.getElementById(objname);
	obj.className="d_on";
	//alert(objname);
	switch (objname){
		case "d_uname":
			strtxt=msg[0];
			break;
		case "d_udomain":
			strtxt=msg[1];
			break;
		case "d_upwd1":
			strtxt=msg[2];
			break;
		case "d_upwd2":
			strtxt=msg[3];
			break;
		case "d_an":
			strtxt=msg[5];
			break;
		case "d_email":
			strtxt=msg[6];
			break;
		case "d_blogname":
			strtxt=msg[7];
			break;
	}
	obj.innerHTML=strtxt;
}

function reset_code(){
	var obj=document.getElementById("ob_codeimg");
	if (obj.tagName=='IMG')
	{		obj.src=obj.src;

	}else{
	obj.onclick;
	}

}
function out_uname(){
	var obj=document.getElementById("d_uname");
	var str=sl(document.getElementById("uname").value);
	var chk=true;
	//alert(str);
	if (str<4 || str>14){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='用户名已经输入。';
		document.getElementById("d_uname_img").src=li_ok;
		if (document.getElementById("blogname").value=='' || !bname_m){
			document.getElementById("blogname").value=document.getElementById("uname").value+"的blog";
			out_blogname();
		}
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[0];
		document.getElementById("d_uname_img").src=li_err;
	}
	return chk;
}

function out_udomain(){
	var obj=document.getElementById("d_udomain");
	var str=document.getElementById("domain").value;
	var chk=true;
	if (str=='' || str.length<4 || str.length>14){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='域名已经输入。';
		document.getElementById("d_udomain_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[1];
		document.getElementById("d_udomain_img").src=li_err;
	}
	return chk;
}
function out_upwd1(){
	var obj=document.getElementById("d_upwd1");
	var str=document.getElementById("upwd").value;
	var chk=true;
	if (str=='' || str.length<6 || str.length>14){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='密码已经输入。';
		document.getElementById("d_upwd1_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[2];
		document.getElementById("d_upwd1_img").src=li_err;
	}
	return chk;
}

function out_upwd2(){
	var obj=document.getElementById("d_upwd2");
	var str=document.getElementById("repassword").value;
	var chk=true;
	if (str!=document.getElementById("upwd").value||str==''){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='重复密码输入正确。';
		document.getElementById("d_upwd2_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[3];
		document.getElementById("d_upwd2_img").src=li_err;
	}
	return chk;
}

function out_question(){
	var obj=document.getElementById("d_question");
	var str=document.getElementById("question").value;
	var chk=true;
	if (str==''){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='密码提示问题已经选择。';
		document.getElementById("d_question_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[4];
		document.getElementById("d_question_img").src=li_err;
	}
	return chk;
}

function out_an(){
	var obj=document.getElementById("d_an");
	var str=sl(document.getElementById("answer").value);
	var chk=true;
	if (str<5 || str>40){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='密码提示问题答案已经输入。';
		document.getElementById("d_an_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[5];
		document.getElementById("d_an_img").src=li_err;
	}
	return chk;
}

function out_email(){
	var obj=document.getElementById("d_email");
	var str=document.getElementById("email").value;
	var chk=true;
	if (str==''|| !str.match(/^[\w\.\-]+@([\w\-]+\.)+[a-z]{2,4}$/ig)){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='电子邮箱地址已经输入。';
		document.getElementById("d_email_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[6];
		document.getElementById("d_email_img").src=li_err;
	}
	return chk;
}

function out_blogname(){
	var obj=document.getElementById("d_blogname");
	var str=document.getElementById("blogname").value;
	var chk=true;
	if (str==''){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='Blog名已经输入。';
		document.getElementById("d_blogname_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[7];
		document.getElementById("d_blogname_img").src=li_err;
	}
	bname_m=true;
	return chk;
}

function out_city(){
	var obj=document.getElementById("d_city");
	var str=document.getElementById("city").value;
	var chk=true;
	if (str==''){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='您所在的地区已经选择。';
		document.getElementById("d_city_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[8];
		document.getElementById("d_city_img").src=li_err;
	}
	return chk;
}

function out_usertype(){
	var obj=document.getElementById("d_usertype");
	var str=document.getElementById("usertype").value;
	var chk=true;
	if (str==0){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='您Blog的类型已经选择。';
		document.getElementById("d_usertype_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[9];
		document.getElementById("d_usertype_img").src=li_err;
	}
	return chk;
}

function out_passregtext(){
	var obj=document.getElementById("d_passregtext");
	var chk=true;

	if (document.oblogform.passregtext[1].checked){chk=false}
	//alert(chk);
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='您已经同意了注册条款。';
		document.getElementById("d_passregtext_img").src=li_ok;
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[10];
		document.getElementById("d_passregtext_img").src=li_err;
	}
	return chk;
}
function chk_reg(){
	var chk=true
	if (!out_uname()){chk=false}
	<%If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then%>
	if (!out_udomain()){chk=false}
	<%end if%>
	if (!out_upwd1()){chk=false}
	if (!out_upwd2()){chk=false}
	if (!out_question()){chk=false}
	if (!out_an()){chk=false}
	if (!out_email()){chk=false}
	if (!out_blogname()){chk=false}
	if (!out_city()){chk=false}
	if (!out_blogname()){chk=false}
	if (!out_usertype()){chk=false}
	if (!out_passregtext()){chk=false}
	if(chk){
	document.getElementById('save_stat').innerHTML='<img src="images/loading.gif" align="absmiddle" />数据提交中……请稍候……'
	document.getElementById('regbotton').disabled='disabled';
	var username=document.oblogform.uname.value;
	var password=document.oblogform.upwd.value;
	var repassword=document.oblogform.repassword.value;
	<%If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then%>
	var domain=document.oblogform.domain.value;
	var domainroot=document.oblogform.user_domainroot.value;
	<%else%>
	var domain='';
	var domainroot='';
	<%end if%>
	<%If oblog.CacheConfig(17)=1 Then%>
		var obcode=document.oblogform.obcode.value;
	<%else%>
		var obcode='';
	<%End if%>
	var question=document.oblogform.question.value;
	var answer=document.oblogform.answer.value;
	var email=document.oblogform.email.value;
	var province=document.oblogform.province.value;
	var city=document.oblogform.city.value;
	var usertype=document.oblogform.usertype.value;
	<%If oblog.CacheConfig(16)=1 Then%>
	var codestr=document.oblogform.codestr.value;
	var ob_codename=document.oblogform.ob_codename.value;
	<%else%>
	var codestr='';
	var ob_codename='';
	<%end if%>
	var blogname=document.oblogform.blogname.value
	var Ajax = new oAjax("reg.asp?action=<%=Application(oblog.Cache_Name & "_RegKey")%>",show_returnsave);
	var arrKey = new Array("username",
							"password",
							"repassword",
							"domain",
							"user_domainroot",
							"question",
							"answer",
							"email",
							"province",
							"city",
							"obcode",
							"codestr",
							"ob_codename",
							"blogname",
							"usertype");
	var arrValue = new Array(username,
							password,
							repassword,
							domain,
							domainroot,
							question,
							answer,
							email,
							province,
							city,
							obcode,
							codestr,
							ob_codename,
							blogname,
							usertype);
	Ajax.Post(arrKey,arrValue);
	//reset_code();
	}
}

function show_returnssn(arrobj){
	if (arrobj){
		var oDialog = new dialog("<%=blogurl%>");
		oDialog.init();
		//alert(arrobj[1]);
		oDialog.set('src',arrobj[1]);
		oDialog.event(arrobj[0],'');
		oDialog.button('dialogOk',"document.getElementById('ssnbotton').disabled=''");
		document.getElementById('chkssn_stat').innerHTML='';
	}
}

function show_returnsave(arrobj){
	var href=''
	if (arrobj){
		switch (arrobj[2]){
			case 'index':
			href="window.location='"+window.location.href.substring(0,window.location.href.lastIndexOf("/"))+"/"+"index.asp'";
			break;
			case 'user_index':
			//href="window.location='"+window.location.href.substring(0,window.location.href.lastIndexOf("/"))+"/"+"user_index.asp?url=user_template.asp?u=new'";
			href="window.location='"+window.location.href.substring(0,window.location.href.lastIndexOf("/"))+"/"+"user_index.asp'";
			break;
		}
		//alert(arrobj[2]);
		//alert(href);
		if(href==''){
		href+="document.getElementById('regbotton').disabled='';";}
		var oDialog = new dialog("<%=blogurl%>");
		oDialog.init();
		oDialog.set('src',arrobj[1]);
		oDialog.event(arrobj[0],'');
		oDialog.button('dialogOk',href);
		document.getElementById('save_stat').innerHTML='';
		if (chkdiv('ob_codeimg'))
		{
			var ob_codeimg=document.getElementById("ob_codeimg");
			if (arrobj[2]!="user_index") ob_codeimg.src=ob_codeimg.src+"&t="+Math.random();
		}
		//document.getElementById('regbotton').disabled='';
	}
}

function sl(st){
	sl1=st.length;
	strLen=0;
	for(i=0;i<sl1;i++){
		if(st.charCodeAt(i)>255) strLen+=2;
	 else strLen++;
	}
	return strLen;
}

function checkssn() {
	var chk=true
	if (!out_uname()){chk=false}
	<%If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then%>
	if (!out_udomain()){chk=false}
	<%end if%>
	if(chk){
		document.getElementById('ssnbotton').disabled='disabled';
		document.getElementById('chkssn_stat').innerHTML='<img src="images/loading.gif" align="absmiddle" />';
		var ssn=document.oblogform.uname.value;
		var email=document.oblogform.email.value;
		<%If oblog.CacheConfig(4)<>"" And oblog.CacheConfig(5) = 1 Then%>
		var domain=document.oblogform.domain.value;
		var domainroot=document.oblogform.user_domainroot.value;
		<%else%>
		var domain='';
		var domainroot='';
		<% End if%>
		var Ajax = new oAjax("reg.asp?action=checkssn",show_returnssn);
		var arrKey = new Array("username", "domain","domainroot","email");
		var arrValue = new Array(ssn,domain,domainroot,email);
		Ajax.Post(arrKey,arrValue);
	}
}

function checkObCode(){
	var obcode=document.oblogform.obcode.value;
	if (obcode==""){
   		alert("邀请码不能为空!");
   		document.oblogform.obcode.focus();
   	}
   	else{
		SendRequest("AjaxServer.asp?action=chkobcode&obcode="+obcode,"msg","");
	}
}


function checkerr(string)
{
var i=0;
for (i=0; i<string.length; i++)
{
if((string.charAt(i) < '0' || string.charAt(i) > '9')  &&  (string.charAt(i) < 'a' || string.charAt(i) > 'z') &&  (string.charAt(i)!='-'))
{
return 1;
}
}
return 0;//pass
}

-->
</script>
<script type="text/javascript">
function GEId(id){return document.getElementById(id);}
function DispPwdStrength(iN,sHL){
	if(iN>3){ iN=3;}
	for(var i=1;i<4;i++){
		var sHCR="ob_pws0";
		if(i<=iN){ sHCR=sHL;}
		if(iN>0){
		GEId("idSM"+i).className=sHCR;
		}
		//GEId("idSMT"+i).className="ob_pwfont2";
		if (iN>0){
			if (i<=iN){
			GEId("idSMT"+i).style.display=((i==iN)?"inline":"none");
			}
		}
		else{
		GEId("idSMT"+i).style.display=((i==iN)?"none":"inline");
		}
	}
}
/*密码强度 来自.Net Passport注册站*/
function EvalPwdStrength(sP){
	if(ClientSideStrongPassword(sP,gSimilarityMap,gDictionary)){
		DispPwdStrength(3,'ob_pws3');
	}else if(ClientSideMediumPassword(sP,gSimilarityMap,gDictionary)){
		DispPwdStrength(2,'ob_pws2');
	}else if(ClientSideWeakPassword(sP,gSimilarityMap,gDictionary)){
		DispPwdStrength(1,'ob_pws1');
	}else{
		DispPwdStrength(0,'ob_pws0');
	}
}


var kNoCanonicalCounterpart = 0;
var kCapitalLetter = 0;
var kSmallLetter = 1;
var kDigit = 2;
var kPunctuation = 3;
var kAlpha =  4;
var kCanonicalizeLettersOnly = true;
var kCananicalizeEverything = false;
var gDebugOutput = null;
var kDebugTraceLevelNone = 0;
var kDebugTraceLevelSuperDetail = 120;
var kDebugTraceLevelRealDetail = 100;
var kDebugTraceLevelAll = 80;
var kDebugTraceLevelMost = 60;
var kDebugTraceLevelFew = 40;
var kDebugTraceLevelRare = 20;
var gDebugTraceLevel = kDebugTraceLevelNone;
function DebugPrint()
{
var string = "";
if (gDebugTraceLevel && gDebugOutput &&
DebugPrint.arguments && (DebugPrint.arguments.length > 1) && (DebugPrint.arguments[0] <= gDebugTraceLevel))
{
for(var index = 1; index < DebugPrint.arguments.length; index++)
{
string += DebugPrint.arguments[index] + " ";
}
string += "<br>\n";
gDebugOutput(string);
}
}
function CSimilarityMap()
{
this.m_elements = "";
this.m_canonicalCounterparts = "";
}
function SimilarityMap_Add(element, canonicalCounterpart)
{
this.m_elements += element;
this.m_canonicalCounterparts += canonicalCounterpart;
}
function SimilarityMap_Lookup(element)
{
var canonicalCounterpart = kNoCanonicalCounterpart;
var index = this.m_elements.indexOf(element);
if (index >= 0)
{
canonicalCounterpart = this.m_canonicalCounterparts.charAt(index);
}
else
{
}
return canonicalCounterpart;
}
function SimilarityMap_GetCount()
{
return this.m_elements.length;
}
CSimilarityMap.prototype.Add = SimilarityMap_Add;
CSimilarityMap.prototype.Lookup = SimilarityMap_Lookup;
CSimilarityMap.prototype.GetCount = SimilarityMap_GetCount;
function CDictionaryEntry(length, wordList)
{
this.m_length = length;
this.m_wordList = wordList;
}
function DictionaryEntry_Lookup(strWord)
{
var fFound = false;
if (strWord.length == this.m_length)
{
var nFirst = 0;
var nLast = this.m_wordList.length - 1;
while( nFirst <= nLast )
{
var nCurrent = Math.floor((nFirst + nLast)/2);
if( strWord == this.m_wordList[nCurrent])
{
fFound = true;
break;
}
else if ( strWord > this.m_wordList[nCurrent])
{
nLast = nCurrent - 1;
}
else
{
nFirst = nCurrent + 1;
}
}
}

return fFound;
}
CDictionaryEntry.prototype.Lookup = DictionaryEntry_Lookup;
function CDictionary()
{
this.m_entries = new Array()
}
function Dictionary_Lookup(strWord)
{
for (var index = 0; index < this.m_entries.length; index++)
{
if (this.m_entries[index].Lookup(strWord))
{
return true;
}
}
}
function Dictionary_Add(length, wordList)
{
var iL=this.m_entries.length;
var cD=new CDictionaryEntry(length, wordList)
this.m_entries[iL]=cD;
}
CDictionary.prototype.Lookup = Dictionary_Lookup;
CDictionary.prototype.Add = Dictionary_Add;
var gSimilarityMap = new CSimilarityMap();
var gDictionary = new CDictionary();
function CharacterSetChecks(type, fResult)
{
this.type = type;
this.fResult = fResult;
}
function isctype(character, type, nDebugLevel)
{
var fResult = false;
switch(type)
{
case kCapitalLetter:
if((character >= 'A') && (character <= 'Z'))
{
fResult = true;
}
break;
case kSmallLetter:
if ((character >= 'a') && (character <= 'z'))
{
fResult = true;
}
break;
case kDigit:
if ((character >= '0') && (character <= '9'))
{
fResult = true;
}
break;
case kPunctuation:
if ("!@#$%^&*()_+-='\";:[{]}\|.>,</?`~".indexOf(character) >= 0)
{
fResult = true;
}
break;
case kAlpha:
if (isctype(character, kCapitalLetter) || isctype(character, kSmallLetter))
{
fResult = true;
}
break;
default:
break;
}

return fResult;
}
function CanonicalizeWord(strWord, similarityMap, fLettersOnly)
{
var canonicalCounterpart = kNoCanonicalCounterpart;
var strCanonicalizedWord = "";
var nStringLength = 0;
if ((strWord != null) && (strWord.length > 0))
{
strCanonicalizedWord = strWord;
strCanonicalizedWord = strCanonicalizedWord.toLowerCase();

if (similarityMap.GetCount() > 0)
{
nStringLength = strCanonicalizedWord.length;

for(var index = 0; index < nStringLength; index++)
{
if (fLettersOnly && !isctype(strCanonicalizedWord.charAt(index), kSmallLetter, kDebugTraceLevelSuperDetail))
{
continue;
}

canonicalCounterpart = similarityMap.Lookup(strCanonicalizedWord.charAt(index));
if (canonicalCounterpart != kNoCanonicalCounterpart)
{
strCanonicalizedWord = strCanonicalizedWord.substring(0, index) + canonicalCounterpart +
strCanonicalizedWord.substring(index + 1, nStringLength);
}
}
}
}
return strCanonicalizedWord;
}
function IsLongEnough(strWord, nAtLeastThisLong)
{
if ((strWord == null) || isNaN(nAtLeastThisLong))
{
return false;
}
else if (strWord.length < nAtLeastThisLong)
{
return false;
}

return true;
}
function SpansEnoughCharacterSets(strWord, nAtLeastThisMany)
{
var nCharSets = 0;
var characterSetChecks = new Array(
new CharacterSetChecks(kCapitalLetter, false),
new CharacterSetChecks(kSmallLetter, false),
new CharacterSetChecks(kDigit, false),
new CharacterSetChecks(kPunctuation, false)
);
if ((strWord == null) || isNaN(nAtLeastThisMany))
{
return false;
}

for(var index = 0; index < strWord.length; index++)
{
for(var nCharSet = 0; nCharSet < characterSetChecks.length;nCharSet++)
{
if (!characterSetChecks[nCharSet].fResult && isctype(strWord.charAt(index), characterSetChecks[nCharSet].type, kDebugTraceLevelAll))
{
characterSetChecks[nCharSet].fResult = true;
break;
}
}
}
for(var nCharSet = 0; nCharSet < characterSetChecks.length;nCharSet++)
{
if (characterSetChecks[nCharSet].fResult)
{
nCharSets++;
}
}

if (nCharSets < nAtLeastThisMany)
{
return false;
}

return true;
}
function FoundInDictionary(strWord, similarityMap, dictionary)
{
var strCanonicalizedWord = "";

if((strWord == null) || (similarityMap == null) || (dictionary == null))
{
return true;
}
strCanonicalizedWord = CanonicalizeWord(strWord, similarityMap, kCanonicalizeLettersOnly);

if (dictionary.Lookup(strCanonicalizedWord))
{
return true;
}

return false;
}
function IsCloseVariationOfAWordInDictionary(strWord, threshold, similarityMap, dictionary)
{
var strCanonicalizedWord = "";
var nMinimumMeaningfulMatchLength = 0;

if((strWord == null) || isNaN(threshold) || (similarityMap == null) || (dictionary == null))
{
return true;
}
strCanonicalizedWord = CanonicalizeWord(strWord, similarityMap, kCananicalizeEverything);
nMinimumMeaningfulMatchLength = Math.floor((threshold) * strCanonicalizedWord.length);
for (var nSubStringLength = strCanonicalizedWord.length; nSubStringLength >= nMinimumMeaningfulMatchLength; nSubStringLength--)
{
for(var nSubStringStart = 0; (nSubStringStart + nMinimumMeaningfulMatchLength) < strCanonicalizedWord.length; nSubStringStart++)
{
var strSubWord = strCanonicalizedWord.substr(nSubStringStart, nSubStringLength);

if (dictionary.Lookup(strSubWord))
{
return true;
}
}
}
return false;
}
function Init()
{
gSimilarityMap.Add('3', 'e');
gSimilarityMap.Add('x', 'k');
gSimilarityMap.Add('5', 's');
gSimilarityMap.Add('$', 's');
gSimilarityMap.Add('6', 'g');
gSimilarityMap.Add('7', 't');
gSimilarityMap.Add('8', 'b');
gSimilarityMap.Add('|', 'l');
gSimilarityMap.Add('9', 'g');
gSimilarityMap.Add('+', 't');
gSimilarityMap.Add('@', 'a');
gSimilarityMap.Add('0', 'o');
gSimilarityMap.Add('1', 'l');
gSimilarityMap.Add('2', 'z');
gSimilarityMap.Add('!', 'i');
gDictionary.Add(3,
"oat|not|ken|keg|ham|hal|gas|cpu|cit|bop|bah".split("|"));
gDictionary.Add(4,
"zeus|ymca|yang|yaco|work|word|wool|will|viva|vito|vita|visa|vent|vain|uucp|util|utah|unix|trek|town|torn|tina|time|tier|tied|tidy|tide|thud|test|tess|tech|tara|tape|tapa|taos|tami|tall|tale|spit|sole|sold|soil|soft|sofa|soap|slav|slat|slap|slam|shit|sean|saud|sash|sara|sand|sail|said|sago|sage|saga|safe|ruth|russ|rusk|rush|ruse|runt|rung|rune|rove|rose|root|rick|rich|rice|reap|ream|rata|rare|ramp|prod|pork|pete|penn|penh|pend|pass|pang|pane|pale|orca|open|olin|olga|oldy|olav|olaf|okra|okay|ohio|oath|numb|null|nude|note|nosy|nose|nita|next|news|ness|nasa|mike|mets|mess|math|mash|mary|mars|mark|mara|mail|maid|mack|lyre|lyra|lyon|lynx|lynn|lucy|love|lose|lori|lois|lock|lisp|lisa|leah|lass|lash|lara|lank|lane|lana|kink|keri|kemp|kelp|keep|keen|kate|karl|june|judy|judo|judd|jody|jill|jean|jane|isis|iowa|inna|holm|help|hast|half|hale|hack|gust|gush|guru|gosh|gory|golf|glee|gina|germ|gatt|gash|gary|game|fred|fowl|ford|flea|flax|flaw|finn|fink|film|fill|file|erin|emit|elmo|easy|done|disk|disc|diet|dial|dawn|dave|data|dana|damn|dame|crab|cozy|coke|city|cite|chem|chat|cats|burl|bred|bill|bilk|bile|bike|beth|beta|benz|beau|bath|bass|bart|bank|bake|bait|bail|aria|anne|anna|andy|alex|abcd".split("|"));
gDictionary.Add(5,
"yacht|xerox|wilma|willy|wendy|wendi|water|warez|vitro|vital|vitae|vista|visor|vicky|venus|venom|value|ultra|u.s.a|tubas|tress|tramp|trait|tracy|traci|toxic|tiger|tidal|thumb|texas|test2|test1|terse|terry|tardy|tappa|tapis|tapir|taper|tanya|tansy|tammy|tamie|taint|sybil|suzie|susie|susan|super|steph|stacy|staci|spark|sonya|sonia|solar|soggy|sofia|smile|slave|slate|slash|slant|slang|simon|shiva|shell|shark|sharc|shack|scrim|screw|scott|scorn|score|scoot|scoop|scold|scoff|saxon|saucy|satan|sasha|sarah|sandy|sable|rural|rupee|runty|runny|runic|runge|rules|ruben|royal|route|rouse|roses|rolex|robyn|robot|robin|ridge|rhode|revel|renee|ranch|rally|radio|quark|quake|quail|power|polly|polis|polio|pluto|plane|pizza|photo|phone|peter|perry|penna|penis|paula|patty|parse|paris|parch|paper|panic|panel|olive|olden|okapi|oasis|oaken|nurse|notre|notch|nancy|nagel|mouse|moose|mogul|modem|merry|megan|mckee|mckay|mcgee|mccoy|marty|marni|mario|maria|marcy|marci|maint|maine|magog|magic|lyric|lyons|lynne|lynch|louis|lorry|loris|lorin|loren|linda|light|lewis|leroy|laura|later|lasso|laser|larry|ladle|kinky|keyes|kerry|kerri|kelly|keith|kazoo|kayla|kathy|karie|karen|julie|julia|joyce|jenny|jenni|japan|janie|janet|james|irene|inane|impel|idaho|horus|horse|honey|honda|holly|hello|heidi|hasty|haste|hamal|halve|haley|hague|hager|hagen|hades|guest|guess|gucci|group|grahm|gouge|gorse|gorky|glean|gleam|glaze|ghoul|ghost|gauss|gauge|gaudy|gator|gases|games|freer|fovea|float|fiona|finny|filly|field|erika|erica|enter|enemy|empty|emily|email|elmer|ellis|ellen|eight|eerie|edwin|edges|eatme|earth|eager|dulce|donor|donna|diane|diana|delay|defoe|david|danny|daisy|cuzco|cubit|cozen|coypu|coyly|cowry|condo|class|cindy|cigar|chess|cathy|carry|carol|carla|caret|caren|candy|candi|burma|burly|burke|brian|breed|borax|booze|booty|bloom|blood|bitch|bilge|bilbo|betty|beryl|becky|beach|bathe|batch|basic|bantu|banks|banjo|baird|baggy|azure|arrow|array|april|anita|angie|amber|amaze|alpha|alisa|alike|align|alice|alias|album|alamo|aires|admin|adept|adele|addle|addis|added|acura|abyss|abcde|1701d|123go|!@#$%".split("|"));
gDictionary.Add(6,
"yankee|yamaha|yakima|y7u8i9|xyzxyz|wombat|wizard|wilson|willie|weenie|warren|visual|virgin|viking|venous|venice|venial|vasant|vagina|ursula|urchin|uranus|uphill|umpire|u.s.a.|tuttle|trisha|trails|tracie|toyota|tomato|toggle|tidbit|thorny|thomas|terror|tennis|taylor|target|tardis|tappet|taoist|tannin|tanner|tanker|tamara|system|surfer|summer|subway|stacie|stacey|spring|sondra|solemn|soleil|solder|solace|soiree|soften|soffit|sodium|sodden|snoopy|snatch|smooch|smiles|slavic|slater|single|singer|simple|sherri|sharon|sharks|sesame|sensor|secret|second|season|search|scroll|scribe|scotty|scooby|schulz|school|scheme|saturn|sandra|sandal|saliva|saigon|sahara|safety|safari|sadism|saddle|sacral|russel|runyon|runway|runoff|runner|ronald|romano|rodent|ripple|riddle|ridden|reveal|return|remote|recess|recent|realty|really|reagan|raster|rascal|random|radish|radial|racoon|racket|racial|rachel|rabbit|qwerty|qawsed|puppet|puneet|public|prince|presto|praise|poster|polite|polish|policy|police|plover|pierre|phrase|photon|philip|persia|peoria|penmen|penman|pencil|peanut|parrot|parent|pardon|papers|pander|pamela|pallet|palace|oxford|outlaw|osiris|orwell|oregon|oracle|olivia|oliver|olefin|office|notion|notify|notice|notate|notary|noreen|nobody|nicole|newton|nevada|mutant|mozart|morley|monica|moguls|minsky|mickey|merlin|memory|mellon|meagan|mcneil|mcleod|mclean|mckeon|mchugh|mcgraw|mcgill|mccann|mccall|mccabe|mayfly|maxine|master|massif|maseru|marvin|markus|malcom|mailer|maiden|magpie|magnum|magnet|maggot|lorenz|lisbon|limpid|leslie|leland|latest|latera|latent|lascar|larkin|langur|landis|landau|lambda|kristy|kristi|krista|knight|kitten|kinney|kerrie|kernel|kermit|kennan|kelvin|kelsey|kelley|keller|keenan|katina|karina|kansas|juggle|judith|jsbach|joshua|joseph|johnny|joanne|joanna|jixian|jimmie|jimbob|jester|jeanne|jasmin|janice|jaguar|jackie|island|invest|instar|ingrid|ingres|impute|holmes|holman|hockey|hidden|hawaii|hasten|harvey|harold|hamlin|hamlet|halite|halide|haggle|haggis|hadron|hadley|hacker|gustav|gusset|gurkha|gurgle|guntis|guitar|graham|gospel|gorton|gorham|gorges|golfer|glassy|ginger|gibson|ghetto|german|george|gauche|gasify|gambol|gamble|gambit|friend|freest|fourth|format|flower|flaxen|flaunt|flakes|finley|finite|fillip|fillet|filler|filled|fermat|fender|fatten|fatima|fathom|father|evelyn|euclid|estate|enzyme|engine|employ|emboss|elanor|elaine|eileen|eighty|eighth|effect|efface|eeyore|eerily|edwina|easier|durkin|durkee|during|durham|duress|duncan|donner|donkey|donate|donald|domino|disney|dieter|device|denise|deluge|delete|debbie|deaden|ddurer|dapper|daniel|dancer|damask|dakota|daemon|cuvier|cuddly|cuddle|cuckoo|cretin|create|cozier|coyote|cowpox|cooper|cookie|connie|coneck|condom|coffee|citrus|citron|citric|circus|charon|change|censor|cement|celtic|cecily|cayuga|catnip|catkin|cation|castle|carson|carrot|carrie|carole|carmen|caress|cantor|burley|burlap|buried|burial|brenda|bremen|breezy|breeze|breech|brandy|brandi|border|borden|borate|bloody|bishop|bilbao|bikini|bigred|betsie|berman|berlin|bedbug|became|beavis|beaver|beauty|beater|batman|bathos|barony|barber|baobab|bantus|banter|bantam|banish|bangui|bangor|bangle|bandit|banana|bakery|bailey|bahama|bagley|badass|aztecs|azsxdc|athena|asylum|arthur|arrest|arrear|arrack|arlene|anvils|answer|angela|andrea|anchor|analog|amazon|amanda|alison|alight|alicia|albino|albert|albeit|albany|alaska|adrian|adelia|adduce|addict|addend|accrue|access|abcdef|abcabc|abc123|a1b2c3|a12345|@#$%^&|7y8u9i|1qw23e|1q2w3e|1p2o3i|1a2b3c|123abc|10sne1|0p9o8i|!@#$%^".split("|"));
gDictionary.Add(7,
"yolanda|wyoming|winston|william|whitney|whiting|whatnot|vitriol|vitrify|vitiate|vitamin|visitor|village|vertigo|vermont|venturi|venture|ventral|venison|valerie|utility|upgrade|unknown|unicorn|unhappy|trivial|torrent|tinfoil|tiffany|tidings|thunder|thistle|theresa|test123|terrify|teleost|tarbell|taproot|tapping|tapioca|tantrum|tantric|tanning|takeoff|swearer|suzanne|susanne|support|success|student|squires|sossina|soldier|sojourn|soignee|sodding|smother|slavish|slavery|slander|shuttle|shivers|shirley|sheldon|shannon|service|seattle|scooter|scissor|science|scholar|scamper|satisfy|sarcasm|salerno|sailing|saguaro|saginaw|sagging|saffron|sabrina|russell|rupture|running|runneth|rosebud|receipt|rebecca|realtor|raleigh|rainbow|quarrel|quality|qualify|pumpkin|protect|program|profile|profess|profane|private|prelude|porsche|politic|playboy|phoenix|persona|persian|perseus|perseid|perplex|penguin|pendant|parapet|panoply|panning|panicle|panicky|pangaea|pandora|palette|pacific|olivier|olduvai|oldster|okinawa|oakwood|nyquist|nursery|numeric|number1|nullify|nucleus|nuclear|notused|nothing|newyork|network|neptune|montana|minimum|michele|michael|merriam|mercury|melissa|mcnulty|mcnally|mcmahon|mckenna|mcguire|mcgrath|mcgowan|mcelroy|mcclure|mcclain|mccarty|mcbride|mcadams|mbabane|mayoral|maurice|marimba|manhole|manager|mammoth|malcolm|malaria|mailbox|magnify|magneto|losable|lorinda|loretta|lorelei|lockout|lioness|limpkin|library|lazarus|lathrop|lateran|lateral|kristin|kristie|kristen|kinsman|kingdom|kennedy|kendall|kellogg|keelson|katrina|jupiter|judaism|judaica|jessica|janeiro|inspire|inspect|insofar|ingress|indiana|include|impetus|imperil|holmium|holmdel|herbert|heather|headmen|headman|harmony|handily|hamburg|halifax|halibut|halfway|haggard|hafnium|hadrian|gustave|gunther|gunshot|gryphon|gosling|goshawk|gorilla|gleason|glacier|ghostly|germane|georgia|geology|gaseous|gascony|gardner|gabriel|freeway|fourier|flowers|florida|fishers|finnish|finland|ferrari|felicia|feather|fatigue|fairway|express|expound|emulate|empress|empower|emitted|emerald|embrace|embower|ellwood|ellison|egghead|durward|durrell|drought|donning|donahue|digital|develop|desiree|default|deborah|damming|cynthia|cyanate|cutworm|cutting|cuddles|cubicle|crystal|coxcomb|cowslip|cowpony|cowpoke|console|conquer|connect|comrade|compton|collins|cluster|claudia|classic|citroen|citrate|citizen|citadel|cistern|christy|chester|charles|charity|celtics|celsius|catlike|cathode|carroll|carrion|careful|carbine|carbide|caraway|caravan|camille|burmese|burgess|bridget|breccia|bradley|bopping|blondie|bilayer|beverly|bernard|bermuda|berlitz|berlioz|beowulf|beloved|because|beatnik|beatles|beatify|bassoon|bartman|baroque|barbara|baptism|banshee|banquet|bannock|banning|bananas|bainite|bailiff|bahrein|bagpipe|baghdad|bagging|bacchus|asshole|arrange|arraign|arragon|arizona|ariadne|annette|animals|anatomy|anatole|amatory|amateur|amadeus|allison|alimony|aliases|algebra|albumin|alberto|alberta|albania|alameda|aladdin|alabama|airport|airpark|airfoil|airflow|airfare|airdrop|adenoma|adenine|address|addison|accrual|acclaim|academy|abcdefg|!@#$%^&".split("|"));
gDictionary.Add(8,
"yosemite|y7u8i9o0|wormwood|woodwind|whistler|whatever|warcraft|vitreous|virginia|veronica|venomous|trombone|transfer|tortoise|tientsin|tideland|ticklish|thailand|testtest|tertiary|terrific|terminal|telegram|tarragon|tapeworm|tapestry|tanzania|tantalus|tantalum|sysadmin|symmetry|sunshine|strangle|startrek|springer|sparrows|somebody|solecism|soldiery|softwood|software|softball|socrates|slatting|slapping|slapdash|slamming|simpsons|serenity|security|schwartz|sanctity|sanctify|samantha|salesman|sailfish|sailboat|sagittal|sagacity|sabotage|rushmore|rosemary|rochelle|robotics|reverend|regional|raindrop|rachelle|qwertyui|qwerasdf|qawsedrf|q1w2e3r4|protozoa|prodding|princess|precious|politics|politico|plymouth|pershing|penitent|penelope|pendulum|patricia|password|passport|paranoia|panorama|panicked|pandemic|pandanus|pakistan|painless|operator|olivetti|oleander|oklahoma|notocord|notebook|notarize|nebraska|napoleon|missouri|michigan|michelle|mesmeric|mercedes|mcmullen|mcmillan|mcknight|mckinney|mckinley|mckesson|mckenzie|mcintyre|mcintosh|mcgregor|mcgovern|mcginnis|mcfadden|mcdowell|mcdonald|mcdaniel|mcconnel|mccauley|mccarthy|mccallum|mayapple|masonite|maryland|marjoram|marinate|marietta|maneuver|mandamus|maledict|maladapt|magnuson|magnolia|magnetic|lyrebird|lymphoma|lorraine|lionking|linoleum|limitate|limerick|laterite|landmass|landmark|landlord|landlady|landhold|landfill|kristine|kirkland|kingston|kimberly|khartoum|keystone|kentucky|keeshond|kathrine|kathleen|jubilant|joystick|jennifer|jacobsen|irishman|interpol|internet|insulate|instinct|instable|insomnia|insolent|insolate|inactive|imperial|iloveyou|illinois|hydrogen|hutchins|homework|hologram|holocene|hibernia|hiawatha|heinlein|hebrides|headlong|headline|headland|hastings|hamilton|halftone|halfback|hagstrom|gunsling|gunpoint|gumption|gorgeous|glaucous|glaucoma|glassine|ginnegan|ghoulish|gertrude|geometry|geometer|garfield|gamesman|gamecock|fungible|function|frighten|freetown|foxglove|fourteen|foursome|forsythe|football|flaxseed|flautist|flatworm|flatware|fidelity|exposure|eternity|enthrone|enthrall|enthalpy|entendre|entangle|engineer|emulsion|emulsify|emporium|employer|employee|employed|emmanuel|elliptic|elephant|einstein|eighteen|duration|donnelly|dominion|dlmhurst|delegate|delaware|december|deadwood|deadlock|deadline|deadhead|danielle|cyanamid|cucumber|cristina|criminal|creosote|creation|cowpunch|couscous|conquest|comrades|computer|comprise|compress|colorado|clusters|citation|charming|cerulean|cenozoic|cemetery|cellular|catskill|cationic|catholic|cathodic|catheter|cascades|carriage|caroline|carolina|carefree|cardinal|burgundy|burglary|bumbling|broadway|breeches|bordello|bordeaux|bilinear|bilabial|bernardo|berliner|berkeley|bedazzle|beaumont|beatrice|beatific|bathrobe|baronial|baritone|bankrupt|banister|bakelite|azsxdcfv|asdfqwer|arkansas|appraise|apposite|anything|angerine|ancestry|ancestor|anatomic|anathema|ambiance|alphabet|albright|albrecht|alberich|albacore|alastair|alacrity|airspace|airplane|airfield|airedale|aircraft|airbrush|airborne|aerobics|adrianna|adelaide|additive|addition|addendum|accouter|academic|academia|abcdefgh|abcd1234|a1b2c3d4|7y8u9i0o|7890yuio|1234qwer|0p9o8i7u|0987poiu|!@#$%^&*".split("|"));
gDictionary.Add(9,
"zimmerman|worldwide|wisconsin|wholesale|vitriolic|ventricle|ventilate|valentine|tidewater|testament|territory|tennessee|telephone|telepathy|teleology|telemetry|telemeter|telegraph|tarantula|tarantara|tangerine|supported|superuser|stuttgart|stratford|stephanie|solemnity|softcover|slaughter|slapstick|signature|sheffield|sarcastic|sanctuary|sagebrush|sagacious|runnymede|rochester|receptive|reception|racketeer|professor|princeton|pondering|politburo|policemen|policeman|persimmon|persevere|persecute|percolate|peninsula|penetrate|pendulous|paralytic|panoramic|panicking|panhandle|oligopoly|oligocene|oligarchy|olfactory|oldenburg|nutrition|nurturant|notorious|notoriety|minnesota|microsoft|mcpherson|mcfarland|mcdougall|mcdonnell|mcdermott|mccracken|mccormick|mcconnell|mccluskey|mcclellan|marijuana|malicious|magnitude|magnetron|magnetite|macintosh|lynchburg|louisiana|lissajous|limousine|limnology|landscape|landowner|kinshasha|kingsbury|kibbutzim|kennecott|jamestown|ironstone|invisible|invention|intuitive|intervene|intersect|inspector|insomniac|insolvent|insoluble|impetuous|imperious|imperfect|holocaust|hollywood|hollyhock|headphone|headlight|headdress|headcount|headboard|happening|hamburger|halverson|gustafson|gunpowder|glasswort|glassware|ghostlike|geometric|gaucherie|freewheel|freethink|freestone|foresight|foolproof|extension|expositor|establish|entertain|employing|emittance|ellsworth|elizabeth|eightieth|eightfold|eiderdown|dusenbury|dusenberg|donaldson|dominique|discovery|desperate|delegable|delectate|decompose|decompile|damnation|cutthroat|crabapple|cornelius|conqueror|connubial|commrades|citizenry|christine|christina|chemistry|cellulose|celluloid|catherine|carryover|burlesque|bloodshot|bloodshed|bloodroot|bloodline|bloodbath|bilingual|bilateral|bijective|bijection|bernadine|berkshire|beethoven|beatitude|bakhtiari|asymptote|asymmetry|apprehend|appraisal|apportion|ancestral|anatomist|alexander|albatross|alabaster|alabamian|adenosine|abcabcabc".split("|"));
gDictionary.Add(10,
"washington|volkswagen|topography|tessellate|temptation|telephonic|telepathic|telemetric|telegraphy|tantamount|superstage|slanderous|salamander|qwertyuiop|polynomial|politician|phrasemake|photometry|photolytic|photolysis|photogenic|phosphorus|phosphoric|persiflage|persephone|perquisite|peninsular|penicillin|penetrable|panjandrum|oligoclase|oligarchic|oldsmobile|nottingham|noticeable|noteworthy|mcnaughton|mclaughlin|mccullough|mcallister|malconduct|maidenhair|limitation|lascivious|landowning|landlubber|landlocked|lamination|khrushchev|juggernaut|irrational|invariable|insouciant|insolvable|incomplete|impervious|impersonal|headmaster|glaswegian|geopolitic|geophysics|fourteenth|foursquare|expressive|expression|expository|exposition|enterprise|eightyfold|eighteenth|effaceable|donnybrook|delectable|decolonize|cuttlefish|cuttlebone|compromise|compressor|comprehend|cellophane|carruthers|california|burlington|burgundian|borderline|borderland|bloodstone|bloodstain|bloodhound|bijouterie|biharmonic|bernardino|beaujolais|basketball|bankruptcy|bangladesh|atmosphere|asymptotic|asymmetric|appreciate|apposition|ambassador|amateurish|alimentary|additional|accomplish|1q2w3e4r5t".split("|"));
gDictionary.Add(11,
"yellowstone|venturesome|territorial|telekinesis|sagittarius|safekeeping|politicking|policewoman|photometric|photography|phosphorous|perseverant|persecutory|persecution|penitential|pandemonium|mississippi|marketplace|magnificent|irremovable|interrogate|institution|inspiration|incompetent|impertinent|impersonate|impermeable|headquarter|hamiltonian|halfhearted|hagiography|geophysical|expressible|emptyhanded|eigenvector|deleterious|decollimate|decolletage|connecticut|comptroller|compressive|compression|catholicism|bloodstream|bakersfield|arrangeable|appreciable|anastomotic|albuquerque".split("|"));
gDictionary.Add(12,
"williamsburg|testamentary|qwerasdfzxcv|q1w2e3r4t5y6|perseverance|pennsylvania|penitentiary|malformation|liquefaction|interstitial|inconclusive|incomputable|incompletion|incompatible|incomparable|imperishable|impenetrable|headquarters|geometrician|ellipsometry|decomposable|decommission|compressible|burglarproof|bloodletting|bilharziasis|asynchronous|asymptomatic|ambidextrous|1q2w3e4r5t6y".split("|"));
gDictionary.Add(13,
"ventriloquist|ventriloquism|poliomyelitis|phosphorylate|oleomargarine|massachusetts|jitterbugging|interpolatory|inconceivable|imperturbable|impermissible|decomposition|comprehensive|comprehension".split("|"));
gDictionary.Add(14,
"slaughterhouse|irreproducible|incompressible|comprehensible|bremsstrahlung".split("|"));
gDictionary.Add(15,
"irreconciliable|instrumentation|incomprehension".split("|"));
gDictionary.Add(16,
"incomprehensible".split("|"));
}

function ClientSideStrongPassword()
{
return (IsLongEnough(ClientSideStrongPassword.arguments[0], "7") &&
SpansEnoughCharacterSets(ClientSideStrongPassword.arguments[0], "3") &&
(!(IsCloseVariationOfAWordInDictionary(ClientSideStrongPassword.arguments[0], "0.6",
ClientSideStrongPassword.arguments[1], ClientSideStrongPassword.arguments[2]))));
}

function ClientSideMediumPassword()
{
return (IsLongEnough(ClientSideMediumPassword.arguments[0], "7") &&
SpansEnoughCharacterSets(ClientSideMediumPassword.arguments[0], "2") &&
(!(FoundInDictionary(ClientSideMediumPassword.arguments[0], ClientSideMediumPassword.arguments[1],
ClientSideMediumPassword.arguments[2]))));
}

function ClientSideWeakPassword()
{
return (IsLongEnough(ClientSideWeakPassword.arguments[0], "6") ||
(!(IsLongEnough(ClientSideWeakPassword.arguments[0], "0"))));
}
</script>
<script language = "javascript"   for = "document"  event = "onkeydown" >
	if (event.keyCode == 13   &&  event.srcElement.type != 'button'  &&  event.srcElement.type != 'submit'  &&  event.srcElement.type != 'reset'  &&  event.srcElement.type != 'textarea'  &&  event.srcElement.type != '')
	chk_reg();
</script>
<SCRIPT language=JavaScript type=text/javascript>
<!--//--><![CDATA[//><!--
function doMenu(MenuName){
 var arrMenus = new Array("showpassregtext");
 for (var i=0; i<arrMenus.length; i++){
  if (MenuName == arrMenus[i]) {
   if(document.getElementById(MenuName).style.display == "block"){
    document.getElementById(arrMenus[i]).style.display = "none";
   }else{
    document.getElementById(MenuName).style.display = "block";
   }
  }else{
   document.getElementById(arrMenus[i]).style.display = "none";
  }
 }
}
//--><!]]>
</SCRIPT>