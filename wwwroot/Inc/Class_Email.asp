<%
'----------------------------------------------------------
'Oblog4.0 �ʼ�����ģ��
'��֧�ָ�������/֧�ֲ���HTML��ʽ/֧��������֤/��֧����������
'����ע����֤/������֤/�һ�����
'Ϊ�˷�ֹЧ��Ӱ�죬ϵͳĬ��Ϊ1���ӷ���һ��
'���й���Ա������ϵͳ����Application�л�ȡ
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
		'��¼Application Last
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
					sRet = "���������JMail.Message ʧ�ܣ����ķ�������֧��JMail���"
					Exit Function
				End If

				'-----------------------------------------------------------------------
				oMail.silent = true '����������󣬷���FALSE��TRUE��ֵj
				oMail.Charset = Email_CharSet '�ʼ������ֱ���Ϊ����
				oMail.ContentType = Email_ContentType '�ʼ��ĸ�ʽΪHTML��ʽ
				oMail.AddRecipient  emailTo '�ʼ��ռ��˵ĵ�ַ
				oMail.From = Email_AdminMail '�����˵�E-MAIL��ַ				
				oMail.MailServerUserName = Email_LoginName '�����ʼ���������¼��
				oMail.MailServerPassword = Email_LoginPwd '��¼����
				oMail.Subject = emailTopic '�ʼ��ı��� 
				oMail.Body = emailBody
				oMail.Priority = 1'�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
				If Err<>0 Then
					sRet = sRet & "����ʧ��!ԭ��2��" & Err.Description
				Else
					oMail.Send Email_SMTP
					oMail.ClearRecipients()
				oMail.Close()
				Set oMail=nothing
					If Err<>0 Then
						sRet = sRet & "����ʧ��!ԭ��1��" & Err.Description
					Else
						sRet = sRet & "���ͳɹ�!1"
					End If
				End If
			Case "1"
				'---------------------------------------
				'CDONTS
				'---------------------------------------
				Set oMail = Server.CreateObject("CDONTS.NewMail")
				If Err<>0 Then
					sRet = "���������CDONTS.NewMail ʧ�ܣ����ķ�������֧�ָ����"
					Exit Function
				End If
				oMail.From = Email_AdminEmail
				oMail.To = emailTo
				oMail.Subject = emailTopic
				oMail.BodyFormat = 0
				oMail.MailFormat = 0
				oMail.Body = emailBody
				If Err<>0 Then
					sRet = sRet & "����ʧ��!ԭ��" & Err.Description
				Else
					oMail.Send
					If Err<>0 Then
						sRet = sRet & "����ʧ��!ԭ��" & Err.Description
					Else
						sRet = sRet & "���ͳɹ�!"
					End If
				End If

			Case "2"
				'---------------------------------------
				'AspEmail
				'---------------------------------------
				Set Obj = Server.CreateObject("Persits.MailSender")
				If Err<>0 Then
					sRet = "���������Persits.MailSender ʧ�ܣ����ķ�������֧��ASPMail���"
					Exit Function
				End If
				oMail.Charset = Email_CharSet
				oMail.IsHTML = True
				oMail.username = Admin_LoginName	'����������Ч���û���
				oMail.password = Admin_LoginPwd	'����������Ч������
				oMail.Priority = 1
				oMail.Host = Admin_SMTP
				'oMail.Port = 25			' �����ѡ.�˿�25��Ĭ��ֵ
				oMail.From = Email_AdminEmail
				oMail.Email_AdminName = Email_AdminName	' �����ѡ
				oMail.AddAddress emailTo,emailTo
				oMail.Subject = emailTopic
				oMail.Body = emailBody
				If Err<>0 Then
					sRet = sRet & "����ʧ��!ԭ��" & Err.Description
				Else
					oMail.Send
					If Err<>0 Then
						sRet = sRet & "����ʧ��!ԭ��" & Err.Description
					Else
						sRet = sRet & "���ͳɹ�!"
					End If
				End If
			Case Else
				sRet="ϵͳδָ���κ��ʼ��������"
		End select
		SendMail=sRet
	End Function
	
	'���͸�����ע���û�
	Public Function SendValidAccountMail(sUserName,sEmail)
		Dim sObCode,sUserId,sUrl,iRet,Sql,rs,sContent
		sObCode=GetGUID
		If Not IsObject(conn) Then link_database
	set rs=Server.CreateObject("adodb.recordset")
		rs.Open "select userid From oblog_user Where useremail='" & sEmail & "' ",conn,1,1
		If rs.RecordCount>1 Then
			ErrMsg="�����ʼ���ַ[" & sEmail & "]��ϵͳ�д��ڶ�������ܽ�����֤!"
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
		sContent=sUserName & " , ����<br/><br/>"
		sContent=sContent & "��л��ע��Ϊ" & blogurl & "�Ļ�Ա������ʸõ�ַ��������ʺ���֤<br/>"
		sUrl=blogurl & "check.asp?user=" & sUserName & "&sn=" & sObCode
		sContent=sContent & "<a href=" & sUrl & " target=_blank>" & sUrl & "</a><br>"
		sContent=sContent & "��������ʼ���Ϊ��ȫ���Ʋ���ֱ�ӷ�����������ַ���뽫�����ַ��������ַ���з��ʣ�<br/>"
		sContent=sContent & sUrl
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Email_AdminName
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Now
		SendValidAccountMail=SendMail(sEmail,sUserName & " ���ã�����֤�����ʺ�",sContent)
	End Function
	
	'���ڲ��ʼ���֤[�����ظ����ʼ���ַ��������֤]
	Public Function SendValidUserMail(sEmail)
		Dim rs,sContent,sUserName,sUserId,sObCode,sUrl,iRet,Sql
		set rs=Server.CreateObject("adodb.recordset")
		rs.Open "select userid,username,isMailValid From oblog_user Where email='" & sEmail & "'",conn,1,3
		If rs.RecordCount>1 Then
			ErrMsg="�����ʼ���ַ[" & sEmail & "]��ϵͳ�д��ڶ�������ܽ�����֤!"
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
		sContent=sUserName & " , ����<br/><br/>"
		sContent=sContent & "Ϊ���ܸ��õ�Ϊ��������ṩ���ʷ���������Ҫ�������ʼ���ַ������֤��<br/>"
		sContent=sContent & "����ʸõ�ַ��������ʼ���֤"
		sUrl=blogurl & "check.asp?user=" & sUserName & "&sn=" & sObCode
		sContent=sContent & "<a href=" & sUrl & " target=_blank>" & sUrl & "</a><br>"
		sContent=sContent & "��������ʼ���Ϊ��ȫ���Ʋ���ֱ�ӷ�����������ַ���뽫�����ַ��������ַ���з��ʣ�<br/>"
		sContent=sContent & sUrl
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Email_AdminName
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Now
		SendValidUserMail=SendMail(sEmail,sUserName & " ���ã��ʼ���Ч����֤",sContent,iRet)
	End Function
	
	'�û���ʧ�������һ�[���ȸ��ʼ���Ҫ�ѱ���֤]
	Public Function SendGetPwdMail(sEmail)
		Dim rs,sContent,sUserName,sUserId,sObCode,sUrl,iRet,Sql
		set rs=Server.CreateObject("adodb.recordset")
		rs.Open "select userid,username,isMailValid From oblog_user Where email='" & sEmail & "'",conn,1,3
		If rs.RecordCount>1 Then
			ErrMsg="�����ʼ���ַ[" & sEmail & "]��ϵͳ�д��ڶ�������ܽ��������һصĺ�������!"
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
		sContent=sUserName & " , ����<br/><br/>"
		sContent=sContent & "��ʹ����" & blogurl & "�������һع���<br/>"
		sContent=sContent & "����ʸõ�ַ������ʾ����������������<br/>"
		sUrl=blogurl & "check.asp?user=" & sUserName & "&sn=" & sObCode
		sContent=sContent & "<a href=" & sUrl & " target=_blank>" & sUrl & "</a><br>"
		sContent=sContent & "��������ʼ���Ϊ��ȫ���Ʋ���ֱ�ӷ�����������ַ���뽫�����ַ��������ַ���з��ʣ�<br/>"
		sContent=sContent & sUrl
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Email_AdminName
		sContent=sContent & "<p>&nbsp;</p>"
		sContent=sContent & Now
		SendGetPwdMail=SendMail(sEmail,sUserName & " ���ã������һ�",sContent,iRet)
	End Function
End Class
%>