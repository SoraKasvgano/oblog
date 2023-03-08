<%
'*********************************************************
'File:			Class_API.asp
'Description:	DPO_API Class For oBlog4.0
'Author:		�о�
'HomePage:		http://www.oblog.cn
'BBS			http://bbs.oblog.cn
'Copyright (C)	2004-2005 oblog.cn All rights reserved.
'LastUpdate:	20060913
'*********************************************************

Class DPO_API_OBLOG
	Private objHttp,XmlDoc,appid,API_Key,strXmlPath,reType,dpo_appid
	Public UserName,PassWord,CookieDate,EMail,Question,Answer,userip,Status,ErrStr,FoundErr
	Public Sex,QQ,MSN,UserStatus,TrueName,Birthday,TelePhone,HomePage,Province,City,address

	Private Sub class_initialize()
		appid="oblog46"
		On Error Resume Next
		Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP"&MsxmlVersion)
		Set XmlDoc =Server.CreateObject("Msxml2.DOMDocument"&MsxmlVersion)
	End Sub
	Private Sub class_terminate()
		On Error Resume Next
		If IsObject(objHttp) Then set objHttp = Nothing
		If IsObject(XmlDoc) Then set XmlDoc = Nothing
	End Sub
	'�ָ������ļ��е�url����ֵ�ֱ��ύ��ÿ��url��
	Public Function ProcessMultiPing(strType)
		Dim i,strUrl
		If strTargetUrls="" Then Exit Function
		For i=0 To UBound(aUrls)
			strUrl=aUrls(i)
			If Left(strUrl,7)="http://" Then
				Call SendPost(strUrl,strType)
			End If
		Next
	End Function
	'��ȡXMLģ���ļ�����ֵΪTrueʱ��������Ϣģ�壬��֮�Ƿ�����Ϣģ��
	Public Sub LoadXmlFile(IsRequest)
		If IsRequest Then
			strXmlPath = Server.MapPath(""&blogdir&"api/Request.xml")
		Else
			strXmlPath = Server.Mappath(""&blogdir&"api/Response.xml")
		End If
		XmlDoc.Load(strXmlPath)
	End Sub
	'Post��Զ���Լ����մ����������
	Private Function SendPost(Url,strType)
		reType=strType
		Dim XMLTemp,strXML
		Dim reMessage
		Dim ajax
		dim API_Timeout
		set ajax=new AjaxXml
		API_Key=MD5(UserName&oblog_Key)
		set XMLTemp = Server.CreateObject("Msxml2.DOMDocument"&MsxmlVersion)
		setNodeValue "username", UserName
		setNodeValue "action", strType
		setNodeValue "syskey", API_Key
		setNodeValue "appid", appid
		select Case strType
			Case "reguser","update"
				setNodeValue "password", PassWord
				setNodeValue "email", EMail
				SetNodeValue "question", Question
				setNodeValue "answer", Answer
				setNodeValue "gender", Sex
				setNodeValue "birthday", Birthday
				setNodeValue "qq", QQ
				setNodeValue "msn", MSN
				setNodeValue "telephone", TelePhone
				setNodeValue "homepage", HomePage
				setNodeValue "userip", userip
				setNodeValue "userstatus", UserStatus
				setNodeValue "province", Province
				setNodeValue "city", city
				setNodeValue "address", address
			Case "login"
				setNodeValue "password", PassWord
				setNodeValue "savecookie", CookieDate
				setNodeValue "userip", userip
			Case "checkname"
				setNodeValue "email", email
			Case Else
		End select
		On Error Resume Next
		API_Timeout=10000
		objHttp.setTimeouts API_Timeout,API_Timeout,API_Timeout*6,API_Timeout*6
		objHttp.Open "POST", Url, False, "", ""
'		objHttp.setRequestHeader "Content-Type", "text/xml"
		objHttp.Send XmlDoc
		If objHttp.readystate<>4 Then
			'AJAX����ע�ᣬ��¼�Լ���֤�û��ķ�����Ϣ
			If reType="reguser" Or reType="checkname" Then
				ajax.re(split(Url&"����Ӧ��Ԥ��״ֵ̬Ϊ4��ʵ��Ϊ"&objHttp.readystate&"��$$$","$$$"))
				Response.End
			Else
				AddErrStr(Url&"����Ӧ��Ԥ��״ֵ̬Ϊ4��ʵ��Ϊ"&objHttp.readystate&"��")
				showErr()
			End if
			Exit Function
		End If
		XMLTemp.Async=True
		XMLTemp.ValidateOnParse=False
		XMLTemp.Load(objHttp.ResponseXML)
		If XMLTemp.parseError.errorCode <> 0 Then
			'AJAX����ע�ᣬ��¼�Լ���֤�û��ķ�����Ϣ
			If reType="reguser"  Or reType="checkname" Then
					ajax.re(split(objHttp.Responsetext&"$$$","$$$"))
					Response.End
			Else
				Response.Write objHttp.Responsetext
				Response.End
'				AddErrStr(XMLTemp.ParseError.ErrorCode)
'				AddErrStr(XMLTemp.ParseError.Reason)
			End if
			Exit Function
		Else
			If XMLTemp.getElementsByTagName("status").item(0).text<>0 Then
				dpo_appid=XMLTemp.getElementsByTagName("appid").item(0).text
				reMessage=XMLTemp.getElementsByTagName("message").item(0).text
				'AJAX����ע�ᣬ��¼�Լ���֤�û��ķ�����Ϣ
				If reType="reguser" Or reType="checkname" Then
					ajax.re(split(dpo_appid &"������ʾ��<br />"&reMessage&"$$$","$$$"))
					Response.End
				Else
					AddErrStr(Replace (reMessage,"<li>",""))
					ShowErr()
				End if
				Exit Function
			End If
		End If
		Set XMLTemp=Nothing
	End Function
	'������Ϣ�������
	Public Function SendResult(status,strMsg)
		setNodeValue "appid", appid
		setNodeValue "status", status
		setNodeValue "message",strMsg
		Response.ContentType = "text/xml"
		Response.Charset = "gb2312"
		Response.Clear
		Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
		Response.Write XmlDoc.documentElement.xml
	End Function
	'��ȡ�û���Ϣ������������
	Public Sub GetUser()
		Call SetNodeValue("username", UserName)
		Call SetNodeValue("email", Email)
		Call SetNodeValue("question", Question)
		Call SetNodeValue("answer", Answer)
		Call SetNodeValue("savecookie", CookieDate)
		Call SetNodeValue("truename", TrueName)
		Call SetNodeValue("gender", Sex)
		Call SetNodeValue("birthday", Birthday)
		Call SetNodeValue("qq", QQ)
		Call SetNodeValue("msn", MSN)
		Call SetNodeValue("telephone", Telephone)
		Call SetNodeValue("homepage", Homepage)
		Call SetNodeValue("userip", UserIP)
		Call SetNodeValue("userstatus", userstatus)
		Call SetNodeValue("province", province)
		Call SetNodeValue("city", city)
		Call SetNodeValue("address",address)
	End Sub
	'����ȡ��XMLģ���еĸ���Ԫ�ظ�ֵ
	Private Function SetNodeValue(strNodeName,strNodeValue)
		If IsNull(strNodeValue) or strNodeValue = "" Then Exit Function
		On Error Resume Next
		XmlDoc.selectSingleNode("//"& strNodeName).text = strNodeValue
		If Err Then
			AddErrStr("д����Ϣ�������������ԣ�")
			showErr()
			Exit Function
		End If
	End Function
	'��������
	Private Sub AddErrStr(Message)
		If ErrStr = "" Then
			ErrStr = dpo_appid &"��ʾ����"& Message
		Else
			ErrStr = ErrStr & "_" & Message
		End If
		FoundErr=True
	End Sub
	'ͬ��һ����
	Private Sub ShowErr()
		If reType<>"checkname" Then
			If ErrStr <> "" Then Response.Redirect ""&blogdir&"err.asp?message=" & ErrStr
		Else
			If ErrStr <> "" Then
				Dim errmsg,errmsg1,i
				errmsg=Split(ErrStr,"_")
				For i=0 to UBound(errmsg)
					If i=0 Then
						errmsg1=errmsg1&"<li>"&errmsg(i)
					Else
						errmsg1=errmsg1&"<br><li>"&errmsg(i)
					End If
				Next
				Response.Write(errmsg1)
			End If
		End If
		FoundErr=True
		ErrStr=Empty
    End Sub
End Class
%>