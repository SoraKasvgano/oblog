<%
'----------------------------------------------------------
'Oblog4.0 邮件发送模块
'不支持附件发送/支持部分HTML格式/支持邮箱认证/不支持批量操作
'用于注册认证/邮箱认证/找回密码
'为了防止效率影响，系统默认为1分钟发送一次
'其中管理员信箱由系统配置Application中获取
'----------------------------------------------------------

Class Oblog_Email

	Private oMail,Email_ContentType,Email_CharSet
	Private  Email_AdminMail,Email_AdminName,Email_SMTP,Email_LoginName,Email_LoginPwd,Email_validateSMTP

	Private Sub Class_Initialize()
		Email_ContentType = "text/html"
		Email_CharSet = "gb2312"
		Email_AdminMail=Application(oblog.Cache_Name & "_Compont")(5)
		Email_AdminName=Application(oblog.Cache_Name & "_Compont")(7)
		Email_SMTP=Application(oblog.Cache_Name & "_Compont")(6)
		Email_LoginName=Application(oblog.Cache_Name & "_Compont")(7)
		Email_LoginPwd=Application(oblog.Cache_Name & "_Compont")(8)
		Email_validateSMTP=Application(oblog.Cache_Name & "_Compont")(10)
		'记录Application Last
	End Sub

	Private Sub Class_Terminate()
	On Error Resume Next
		If Isobject(oMail) Then
			Set oMail = Nothing
		End If
	End Sub
	
	
	Public Function SendMail(emailTo,emailTopic,emailBody)
		Dim sRet
	'On Error Resume Next
		select Case Application(oblog.Cache_Name & "_Compont")(4)
			Case "0"
				'---------------------------------------
				'JMail4.x
				'---------------------------------------
				Set oMail = Server.CreateObject("JMail.Message")
				If Err<>0 Then
					sRet = "创建组件：JMail.Message 失败，您的服务器不支持JMail组件"
					Exit Function
				End If

				'-----------------------------------------------------------------------
				oMail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值j
				oMail.Charset = Email_CharSet '邮件的文字编码为国标
				oMail.ContentType = Email_ContentType '邮件的格式为HTML格式
				oMail.AddRecipient  emailTo '邮件收件人的地址
				oMail.From = Email_AdminMail '发件人的E-MAIL地址				
				oMail.MailServerUserName = Email_LoginName '您的邮件服务器登录名
				oMail.MailServerPassword = Email_LoginPwd '登录密码
				oMail.Subject = emailTopic '邮件的标题 
				oMail.Body = emailBody
				oMail.Priority = 1'邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
				If Err<>0 Then
					sRet = sRet & "发送失败!原因2：" & Err.Description
				Else
					oMail.Send Email_SMTP
					oMail.ClearRecipients()
				oMail.Close()
				Set oMail=nothing
					If Err<>0 Then
						sRet = sRet & "发送失败!原因1：" & Err.Description
					Else
						sRet = sRet & "发送成功!1"
					End If
				End If
			Case "1"
				'---------------------------------------
				'CDONTS
				'---------------------------------------
				Set oMail = Server.CreateObject("CDONTS.NewMail")
				If Err<>0 Then
					sRet = "创建组件：CDONTS.NewMail 失败，您的服务器不支持该组件"
					Exit Function
				End If
				oMail.From = Email_AdminEmail
				oMail.To = emailTo
				oMail.Subject = emailTopic
				oMail.BodyFormat = 0
				oMail.MailFormat = 0
				oMail.Body = emailBody
				If Err<>0 Then
					sRet = sRet & "发送失败!原因：" & Err.Description
				Else
					oMail.Send
					If Err<>0 Then
						sRet = sRet & "发送失败!原因：" & Err.Description
					Else
						sRet = sRet & "发送成功!"
					End If
				End If

			Case "2"
				'---------------------------------------
				'AspEmail
				'---------------------------------------
				Set Obj = Server.CreateObject("Persits.MailSender")
				If Err<>0 Then
					sRet = "创建组件：Persits.MailSender 失败，您的服务器不支持ASPMail组件"
					Exit Function
				End If
				oMail.Charset = Email_CharSet
				oMail.IsHTML = True
				oMail.username = Admin_LoginName	'服务器上有效的用户名
				oMail.password = Admin_LoginPwd	'服务器上有效的密码
				oMail.Priority = 1
				oMail.Host = Admin_SMTP
				'oMail.Port = 25			' 该项可选.端口25是默认值
				oMail.From = Email_AdminEmail
				oMail.Email_AdminName = Email_AdminName	' 该项可选
				oMail.AddAddress emailTo,emailTo
				oMail.Subject = emailTopic
				oMail.Body = emailBody
				If Err<>0 Then
					sRet = sRet & "发送失败!原因：" & Err.Description
				Else
					oMail.Send
					If Err<>0 Then
						sRet = sRet & "发送失败!原因：" & Err.Description
					Else
						sRet = sRet & "发送成功!"
					End If
				End If
			Case Else
				sRet="系统未指定任何邮件发送组件"
		End select
		SendMail=sRet
	End Function
	
	'发送给初次注册用户
	Public Function SendValidAccountMail(sUserName,sEmail)
		Dim sObCode,sUserId,sUrl,iRet,Sql,rs,sContent
		sObCode=GetGUID
		If Not IsObject(conn) Then link_database
	set rs=Server.CreateObject("adodb.recordset")
		rs.Open "select userid From oblog_user Where useremail='" & sEmail & "' ",conn,1,1
		If rs.RecordCount>1 Then
			ErrMsg="您的邮件地址[" & sEmail & "]在系统中存在多个，不能进行验证!"
			Set rs=Nothing
			Exit Function
		End If
		sUserId=rs(0)
		rs.Close
		Set rs=Nothing
		'sUserId=oblog.Execute("select top 1 userid From oblog_user Where useremail='" & sEmail & "'")(0)
		Sql="Insert Into oblog_obcodes(obcode,creatuser,createtime,creatip,itype,istate) Values('"
		Sql= Sql &  sObcode & "'," & sUserId &",'" & Now & "','" & oblog.UserIp & "',2,0)"
		oblog.Execute Sql
		sContent=sUserName & " , 您好<br/><br/>"
		sContent=sContent & "感谢您注册为" & blogurl & "的会员，请访问该地址完成您的帐号验证<br/>"
		sUrl=blogurl & "check.asp?user=" & sUserName & "&sn=" & sObCode
		sContent=sContent & "<a href=" & sUrl & " target=_blank>" & sUrl & "</a><br>"
		sContent=sContent & "如果您的邮件因为安全限制不能直接访问呢上述地址，请将下面地址拷贝到地址栏中访问：<br/>"
		sContent=sContent & sUrl
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Email_AdminName
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Now
		SendValidAccountMail=SendMail(sEmail,sUserName & " 您好，请验证您的帐号",sContent)
	End Function
	
	'后期补邮件验证[对于重复的邮件地址不进行验证]
	Public Function SendValidUserMail(sEmail)
		Dim rs,sContent,sUserName,sUserId,sObCode,sUrl,iRet,Sql
		set rs=Server.CreateObject("adodb.recordset")
		rs.Open "select userid,username,isMailValid From oblog_user Where email='" & sEmail & "'",conn,1,3
		If rs.RecordCount>1 Then
			ErrMsg="您的邮件地址[" & sEmail & "]在系统中存在多个，不能进行验证!"
			Set rs=Nothing
			Exit Function
		End If
		sUserId=rs(0)
		sUserName=rs(1)
		rs(2)=1
		rs.Update
		Set rs=Nothing
		sObCode=GetGuid
		oblog.Execute
		Sql="Insert Into oblog_obcodes(obcode,creatuser,createtime,creatip,itype,istate) Values('"
		Sql= Sql &  sObcode & "',"& sUserId &",'" & Now & "','" & oblog.UserIp & "',2,0)"
		oblog.Execute Sql
		sContent=sUserName & " , 您好<br/><br/>"
		sContent=sContent & "为了能更好的为广大网友提供优质服务，我们需要对您的邮件地址进行验证。<br/>"
		sContent=sContent & "请访问该地址完成您的邮件验证"
		sUrl=blogurl & "check.asp?user=" & sUserName & "&sn=" & sObCode
		sContent=sContent & "<a href=" & sUrl & " target=_blank>" & sUrl & "</a><br>"
		sContent=sContent & "如果您的邮件因为安全限制不能直接访问呢上述地址，请将下面地址拷贝到地址栏中访问：<br/>"
		sContent=sContent & sUrl
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Email_AdminName
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Now
		SendValidUserMail=SendMail(sEmail,sUserName & " 您好，邮件有效性验证",sContent,iRet)
	End Function
	
	'用户丢失密码后的找回[首先该邮件需要已被验证]
	Public Function SendGetPwdMail(sEmail)
		Dim rs,sContent,sUserName,sUserId,sObCode,sUrl,iRet,Sql
		set rs=Server.CreateObject("adodb.recordset")
		rs.Open "select userid,username,isMailValid From oblog_user Where email='" & sEmail & "'",conn,1,3
		If rs.RecordCount>1 Then
			ErrMsg="您的邮件地址[" & sEmail & "]在系统中存在多个，不能进行密码找回的后续操作!"
			Set rs=Nothing
			Exit Function
		End If
		sUserId=rs(0)
		sUserName=rs(1)
		rs(2)=1
		rs.Update
		Set rs=Nothing
		sObCode=GetGuid
		oblog.Execute
		Sql="Insert Into oblog_obcodes(obcode,creatuser,createtime,creatip,itype,istate) Values('"
		Sql= Sql &  sObcode & "',"& sUserId &",'" & Now & "','" & oblog.UserIp & "',3,0)"
		oblog.Execute Sql
		sContent=sUserName & " , 您好<br/><br/>"
		sContent=sContent & "您使用了" & blogurl & "的密码找回功能<br/>"
		sContent=sContent & "请访问该地址依照提示重新设置您的密码<br/>"
		sUrl=blogurl & "check.asp?user=" & sUserName & "&sn=" & sObCode
		sContent=sContent & "<a href=" & sUrl & " target=_blank>" & sUrl & "</a><br>"
		sContent=sContent & "如果您的邮件因为安全限制不能直接访问呢上述地址，请将下面地址拷贝到地址栏中访问：<br/>"
		sContent=sContent & sUrl
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Email_AdminName
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Now
		SendGetPwdMail=SendMail(sEmail,sUserName & " 您好，密码找回",sContent,iRet)
	End Function
End Class
%>