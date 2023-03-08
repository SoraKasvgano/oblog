<!-- #include file="bloginterface.asp" -->
<% Response.Buffer=True %>
<%
Dim objUser
Dim GeneralCategories()
Dim SystemCategories()

Function ParseDateForRFC3339(dtmDate)

	Dim dtmDay, dtmWeekDay, dtmMonth, dtmYear
	Dim dtmHours, dtmMinutes, dtmSeconds

	Dim strTimeZone

	dtmYear = Year(dtmDate)
	dtmMonth = Right("00" & Month(dtmDate),2)
	dtmDay = Right("00" & Day(dtmDate),2)

	dtmHours = Right("00" & Hour(dtmDate),2)
	dtmMinutes = Right("00" & Minute(dtmDate),2)
	dtmSeconds = Right("00" & Second(dtmDate),2)

	strTimeZone=Left("+0800",3) & ":" & Right("+0800",2)

	ParseDateForRFC3339 = dtmYear & "-" & dtmMonth & "-" & dtmDay & "T" & dtmHours & ":" & dtmMinutes & ":" & dtmSeconds & strTimeZone

End Function 


Function ReBuildBlog(articleId)
	Dim blog
	set blog=new class_blog
	blog.userid=objUser.Id
	blog.update_log articleId, 0
	
	blog.update_log 0, 0
		
	blog.update_calendar(articleId)
	blog.update_newblog(objUser.Id)
	blog.update_subject(objUser.Id)
	blog.update_index(0)
	
	blog.CreateFunctionPage()
	set blog=nothing
End Function


Function VerifyUser(userName, userPassWord)

	Set objUser = Nothing
	Set objUser = New BlogUser

	objUser.Name = userName
	objUser.PassWord = md5(userPassWord)
	
	If objUser.Verify() Then
		VerifyUser = True
	Else
		Call ResponseError(ERROR_NOT_LEGAL_USER)
	End If

End Function

Function GetUsersBlogs()

	Dim strXML
	strXML="<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><params><param><value><array><data><value><struct><member><name>url</name><value><string>$1</string></value></member><member><name>blogid</name><value><string>$2</string></value></member><member><name>blogName</name><value><string>$3</string></value></member></struct></value></data></array></value></param></params></methodResponse>"

	strXML=Replace(strXML,"$1",TransferHTML("www.oblog.cn","[<][>][&][""]"))
	strXML=Replace(strXML,"$2",TransferHTML(objUser.Id,"[<][>][&][""]"))
	strXML=Replace(strXML,"$3",TransferHTML(objUser.Name,"[<][>][&][""]"))

	Response.Write strXML

End Function

Function GetCategories()

	Dim strXML
	Dim strCategoryInfo

	strXML="<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><params><param><value><array><data><value>$1</value></data></array></value></param></params></methodResponse>"

	strCategoryInfo="<struct><member><name>description</name><value><string>$1</string></value></member><member><name>httpUrl</name><value><string>$2</string></value></member><member><name>rssUrl</name><value><string>$3</string></value></member><member><name>title</name><value><string>$4</string></value></member><member><name>categoryid</name><value><string>$5</string></value></member></struct>"

	GetGeneralCategories()'获取综合分类
	
	Dim Cate
	Dim s
	Dim strCategories
	For Each Cate in GeneralCategories
		If IsObject(Cate) Then
			s=strCategoryInfo
			s=Replace(s,"$1",TransferHTML(Cate.Intro,"[<][>][&][""]"))
			s=Replace(s,"$2",TransferHTML(Cate.Order,"[<][>][&][""]"))
			s=Replace(s,"$3",TransferHTML(Cate.Count,"[<][>][&][""]"))
			s=Replace(s,"$4",TransferHTML(Cate.Name,"[<][>][&][""]"))
			s=Replace(s,"$5",TransferHTML(Cate.Id,"[<][>][&][""]"))
			strCategories=strCategories & s
		End If
	Next

	strXML=Replace(strXML,"$1",strCategories)

	Response.Write strXML

End Function


Function GetCategories2()

	Dim strXML
	Dim strCategoryInfo

	strXML="<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><params><param><value><array><data><value>$1</value></data></array></value></param></params></methodResponse>"

	strCategoryInfo="<struct><member><name>description</name><value><string>$1</string></value></member><member><name>httpUrl</name><value><string>$2</string></value></member><member><name>rssUrl</name><value><string>$3</string></value></member><member><name>title</name><value><string>$4</string></value></member><member><name>categoryid</name><value><string>$5</string></value></member></struct>"

	GetSystemCategories()'获取系统分类
	
	Dim Cate
	Dim s
	Dim strCategories
	For Each Cate in SystemCategories
		If IsObject(Cate) Then
			s=strCategoryInfo
			s=Replace(s,"$1",TransferHTML(Cate.Intro,"[<][>][&][""]"))
			s=Replace(s,"$2",TransferHTML(Cate.Order,"[<][>][&][""]"))
			s=Replace(s,"$3",TransferHTML(Cate.Count,"[<][>][&][""]"))
			s=Replace(s,"$4",TransferHTML(Cate.Name,"[<][>][&][""]"))
			s=Replace(s,"$5",TransferHTML(Cate.Id,"[<][>][&][""]"))
			strCategories=strCategories & s
		End If
	Next

	strXML=Replace(strXML,"$1",strCategories)

	Response.Write strXML

End Function


Function NewPost(structPost, bolPublish)

	Dim objXmlFile
	Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")

	objXmlFile.loadXML(structPost)

	Dim strXML

	strXML = "<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><params><param><value><string>$1</string></value></param></params></methodResponse>"


	Dim objArticle
	Set objArticle = New BlogArticle
	
	objArticle.AuthorId = objUser.Id
	objArticle.Author = objUser.Name
	objArticle.UserId = objUser.Id
	
	If (bolPublish = True) Then
		objArticle.IsDraft = 0'发布
	Else
		objArticle.IsDraft = 1'草稿
	End If
	
	objArticle.Topic = objXmlFile.documentElement.selectSingleNode("member[name=""title""]/value/string").text

	Dim strCate
	strCate = objXmlFile.documentElement.selectSingleNode("member[name=""categories""]/value/array/data/value[0]/string").text
	
	GetGeneralCategories()
	Dim Cate
	For Each Cate in GeneralCategories
		If IsObject(Cate) Then
			If strCate = Cate.Name Then
				objArticle.SubjectId = Cate.Id
				Exit For
			End If
		End If
	Next

	Dim objNode
	Set objNode = Nothing
	
	Set objNode = objXmlFile.documentElement.selectSingleNode("member[name=""categories2""]/value/array/data/value[0]/string")
	If objNode is Nothing  Then
		objArticle.ClassId = 0
	Else
		strCate = objNode.text

		GetSystemCategories()
		For Each Cate in SystemCategories
			If IsObject(Cate) Then
				If strCate = Cate.Name Then
					objArticle.ClassId = Cate.Id
					Exit For
				End If
			End If
		Next
	End If
	
	Set objNode = Nothing
	Set objNode = objXmlFile.documentElement.selectSingleNode("member[name=""pubDate""]/value/string")
	If objNode is Nothing  Then
		objArticle.AddTime = now()
	Else
		objArticle.AddTime = objNode.text	
	End If
	
	Set objNode = Nothing
	Set objNode = objXmlFile.documentElement.selectSingleNode("member[name=""tags""]/value/string")
	If objNode is Nothing  Then
		objArticle.Tags = ""
	Else
		objArticle.Tags = objNode.text	
	End If
	
	Set objNode = Nothing
	Set objNode = objXmlFile.documentElement.selectSingleNode("member[name=""trackback""]/value/string")
	If objNode is Nothing  Then
		objArticle.TrackBack = ""
	Else
		objArticle.TrackBack = objNode.text	
	End If
	
	Set objNode = Nothing
	Set objNode = objXmlFile.documentElement.selectSingleNode("member[name=""abstract""]/value/string")
	If objNode is Nothing  Then
		objArticle.Abstract = ""
	Else
		objArticle.Abstract = objNode.text	
	End If
	

	objArticle.Log_Text = objXmlFile.documentElement.selectSingleNode("member[name=""description""]/value/string").text	
	
	If objArticle.AddNew() = True Then'加入文章
		Call ReBuildBlog(objArticle.Id)
		
		Response.Clear
		strXML = Replace(strXML,"$1",objArticle.ID )
		Response.Write strXML
	Else
		Call ResponseError(ERROR_ACCESS_DATABASE_FAILED)
	End If

End Function


Function GetRecentPosts(numberOfPosts)
	On Error Resume Next
	'判断是否超出允许获取的最多文章数
	If ((CInt(numberOfPosts) > MAX_GETRECENTPOSTS_NUM) And (MAX_GETRECENTPOSTS_NUM <> 0)) Then
		Call ResponseError(ERROR_NOT_LEGAL_GETRECENTPOSTS_NUM)
	End If

	Dim strXML
	Dim strPost
	Dim strRecentPosts

	strXML = "<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><params><param><value><array><data><value>$1</value></data></array></value></param></params></methodResponse>"

	strPost = "<struct><member><name>title</name><value><string>$1</string></value></member><member><name>description</name><value><string>$2</string></value></member><member><name>dateCreated</name><value><dateTime.iso8601>$3</dateTime.iso8601></value></member><member><name>categories</name><value><array><data><value><string>$4</string></value></data></array></value></member><member><name>postid</name><value><string>$5</string></value></member><member><name>userid</name><value><string>$6</string></value></member><member><name>link</name><value><string>$7</string></value></member></struct>"

	Dim s
	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim Cate
	Dim strCate

	Set objRS = Server.CreateObject("ADODB.Recordset")

	objRS.Open "select TOP " & numberOfPosts & " * FROM [oblog_log] WHERE [isdel] = 0 AND [logtype] = 0 AND [authorid] = "& objUser.Id &" ORDER BY [istop] DESC , [logid] DESC",conn,1,1
	If (Not objRS.bof) And (Not objRS.eof) Then

		GetGeneralCategories()

		For i = 1 To objRS.RecordCount
			s=strPost
			Set Cate=Nothing
			For Each Cate in GeneralCategories
				If IsObject(Cate) Then
					If objRS("subjectid") = Cate.Id Then
							strCate = Cate.Name
						Exit For
					End If
				End If
			Next
			s=Replace(s,"$4",TransferHTML(strCate,"[<][>][&][""]"))
			s=Replace(s,"$3",TransferHTML(ParseDateForRFC3339(objRS("addtime")),"[<][>][&][""]"))
			s=Replace(s,"$5",TransferHTML(objRS("logid"),"[<][>][&][""]"))
			s=Replace(s,"$6",TransferHTML(objRS("authorid"),"[<][>][&][""]"))
			s=Replace(s,"$7",TransferHTML(objRS("trackbacknum"),"[<][>][&][""]"))
			s=Replace(s,"$1",TransferHTML(objRS("topic"),"[<][>][&][""]"))
			s=Replace(s,"$2",TransferHTML(objRS("logtext"),"[<][>][&][""]"))'放最后,降低日志中有$n的冲突

			strRecentPosts=strRecentPosts & s

			objRS.MoveNext
		Next

	End If

	objRS.Close
	Set objRS=Nothing

	strXML=Replace(strXML,"$1",strRecentPosts)

	Response.Write strXML

End Function


Function EditPost(intPostID, structPost, bolPublish)

	On Error Resume Next

	Dim objXmlFile
	Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")

	objXmlFile.loadXML(structPost)

	Dim strXML

	strXML="<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><params><param><value><boolean>$1</boolean></value></param></params></methodResponse>"

	Dim objArticle
	Set objArticle=New BlogArticle

	If Not(objArticle.VerifyId(intPostID)) Then
		Call ResponseError(ERROR_NOT_EXIST_ARTICLE)
	End If

	objArticle.Topic = objXmlFile.documentElement.selectSingleNode("member[name=""title""]/value/string").text

	If (bolPublish = True) Then
		objArticle.IsDraft = 0'发布
	Else
		objArticle.IsDraft = 1'草稿
	End If

	GetGeneralCategories()
	Dim strCate
	strCate = objXmlFile.documentElement.selectSingleNode("member[name=""categories""]/value/array/data/value[0]/string").text
	If strCate<>"" Then
		Dim Cate
		For Each Cate in GeneralCategories
			If IsObject(Cate) Then
				If strCate = Cate.Name Then
					objArticle.Subjectid = Cate.Id
					Exit For
				End If
			End If
		Next
	End If
	
	objArticle.Log_Text = objXmlFile.documentElement.selectSingleNode("member[name=""description""]/value/string").text

	objArticle.AuthorId = objUser.Id
	If objArticle.Modify() = True Then
		Call ReBuildBlog(objArticle.Id)

		Response.Clear

		strXML=Replace(strXML,"$1",1)
		Response.Write strXML
	Else
		Call ResponseError(ERROR_ACCESS_DATABASE_FAILED)
	End If

End Function


Function DeletePost(intPostID)

	Dim strXML

	strXML="<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><params><param><value><boolean>$1</boolean></value></param></params></methodResponse>"

	Dim objArticle
	Set objArticle=New BlogArticle

	If Not (objArticle.VerifyId(intPostID)) Then
		Call RespondError(9,ZVA_ErrorMsg(9))
	End If

	If objArticle.Delete() Then
		Call ReBuildBlog(objArticle.Id)
		
		Response.Clear

		strXML=Replace(strXML,"$1",1)
		Response.Write strXML
	Else
		Call RespondseError(11)
	End If

End Function


Function NewMediaObject(strFileName,strFileType,strFileBits)
	'判断是否开启上传文件
	If Not UPLOADFILE_SIGN Then
		Call ResponseError(ERROR_SHUTDOWN_UPLOADFILE)
	End If
	
	Dim strXML
	strXML="<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><params><param><value><struct><member><name>url</name><value><string>$1</string></value></member></struct></value></param></params></methodResponse>"

	Dim objUpLoadFile
	Set objUpLoadFile = New BlogUpLoadFile
	objUpLoadFile.AuthorID = objUser.Id
	objUpLoadFile.FileName = strFileName
	objUpLoadFile.UploadType="Stream"

	Dim xmlnode
	Set xmlnode = objXmlFile.createElement("file")
	xmlnode.datatype = "bin.base64"
	xmlnode.text = strFileBits

	Dim objStreamUp
	Set objStreamUp = Server.CreateObject(oblog.CacheCompont(2))

	With objStreamUp
		.Type = 1
		.Mode = 3
		.Open
		.Position = 0
		.Write xmlnode.nodeTypedvalue
		.Position = 0
		objUpLoadFile.Stream=.Read
		.Close
	End With

	If objUpLoadFile.UpLoad() Then
	
		strXML=Replace(strXML,"$1",TransferHTML(objUpLoadFile.BackUrl,"[<][>][&][""]"))
		Response.Write strXML

	End If

End Function
'/////////////////////////////////////////////////////////////////////////////////////////

Response.ContentType = "text/xml"
Session.CodePage = 936'!GB2312

	'If VerifyUser("nicetoyou","234561") Then Call DeletePost(30)
	'Response.End

Dim strXmlCall
Dim objXmlFile

strXmlCall=Request.BinaryRead(Request.TotalBytes)
Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")

objXmlFile.load(strXmlCall)

If objXmlFile.readyState=4 Then
	If objXmlFile.parseError.errorCode <> 0 Then
		Call ResponseError(ERROR_NOT_LEGAL_XMLREQUEST)
	Else
		If oblog.chkiplock() Then
			Call ResponseError(ERROR_LOCKIP)
			Response.End
		End If

		Dim objRootNode
		Set objRootNode=objXmlFile.documentElement

		Dim strAction
		strAction=objRootNode.selectSingleNode("methodName").text

		Dim strUserName
		Dim strUserPassWord
		Dim intNumberOfPosts
		Dim strPost
		Dim intPostID
		Dim strFileName
		Dim strFileType
		Dim strFileBits
		Dim bolPublish

		select Case strAction
			Case "blogger.getUsersBlogs":
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				
				If VerifyUser(strUserName,strUserPassWord) Then Call GetUsersBlogs()
				
			Case "metaWeblog.getCategories":
				strUserName		=	objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord	=	objRootNode.selectSingleNode("params/param[2]/value/string").text
				
				If VerifyUser(strUserName,strUserPassWord) Then Call GetCategories()
				
			Case "metaWeblog.getCategories2":'此方法为oblog扩展,获取系统分类
				strUserName		=	objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord	=	objRootNode.selectSingleNode("params/param[2]/value/string").text
				
				If VerifyUser(strUserName,strUserPassWord) Then Call GetCategories2()
				
			Case "metaWeblog.newPost":
				strUserName		=	objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord	=	objRootNode.selectSingleNode("params/param[2]/value/string").text
				strPost			=	objRootNode.selectSingleNode("params/param[3]/value/struct").xml
				bolPublish		=	CBool(objRootNode.selectSingleNode("params/param[4]/value/boolean").text)
				
				If VerifyUser(strUserName,strUserPassWord) Then 
					If oblog.CheckPostAccess <> "" Then 
						Call ResponseError(ERROR_GROUP_ISPOSTMAX)
					Else
						Call NewPost(strPost,bolPublish)
					End If
				End if
				
			Case "metaWeblog.getRecentPosts":
				strUserName		=	objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord	=	objRootNode.selectSingleNode("params/param[2]/value/string").text
				intNumberOfPosts=	objRootNode.selectSingleNode("params/param[3]/value/int").text
				
				If VerifyUser(strUserName,strUserPassWord) Then Call GetRecentPosts(intNumberOfPosts)
				
			Case "metaWeblog.editPost":
				intPostID		=	objRootNode.selectSingleNode("params/param[0]/value/string").text
				strUserName		=	objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord	=	objRootNode.selectSingleNode("params/param[2]/value/string").text
				strPost			=	objRootNode.selectSingleNode("params/param[3]/value/struct").xml
				bolPublish		=	CBool(objRootNode.selectSingleNode("params/param[4]/value/boolean").text)
				
				If VerifyUser(strUserName,strUserPassWord) Then Call EditPost(intPostID,strPost,bolPublish)
				
			Case "blogger.deletePost":
				intPostID		=	objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserName		=	objRootNode.selectSingleNode("params/param[2]/value/string").text
				strUserPassWord	=	objRootNode.selectSingleNode("params/param[3]/value/string").text
				
				If VerifyUser(strUserName,strUserPassWord) Then Call DeletePost(intPostID)
				
			Case "metaWeblog.newMediaObject":
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				strFileName=objRootNode.selectSingleNode("params/param[3]/value/struct/member[name=""name""]/value/string").text
				strFileType=objRootNode.selectSingleNode("params/param[3]/value/struct/member[name=""type""]/value/string").text
				strFileBits=objRootNode.selectSingleNode("params/param[3]/value/struct/member[name=""bits""]/value/base64").text
				
				If VerifyUser(strUserName,strUserPassWord) Then Call NewMediaObject(strFileName,strFileType,strFileBits)

			Case Else
				Call ResponseError(ERROR_UNKNOW_BLOGAPIMETHOD)
		End select 

	End If
End If


conn.Close
Set conn = Nothing
%>
