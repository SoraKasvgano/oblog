<!--#include file="../inc/inc_syssite.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<!--#include file="../inc/class_Trackback.asp"-->
<%
Dim afxDebug
Const MAX_GETRECENTPOSTS_NUM				= 0		'getRecentPosts最多允许的文章数量,0为不限制
Const MAX_PUBLISHSPACE_TIME					= 10		'两次发布文章最小时间间隔 单位/秒, 0为不限制
'无效Const MAX_UPLOADFILESPACE_TIME				= 0		'两次上传文件最小时间间隔 单位/秒, 0为不限制

Const UPLOADFILE_SIGN						= true	'是否允许上传文件.

Const ERROR_NOT_LEGAL_XMLREQUEST			= 1		'不是有效格式的XML请求
Const ERROR_UNKNOW_BLOGAPIMETHOD			= 2		'未知的BlogAPI方法
Const ERROR_NOT_LEGAL_USER					= 3		'用户名或密码错误
Const ERROR_NOT_EXIST_ARTICLE				= 4		'要修改的文章不存在
Const ERROR_ACCESS_DATABASE_FAILED			= 5		'ASP端数据库操作失败
Const ERROR_NOT_LEGAL_TITLE					= 6		'标题为空或大于100
Const ERROR_NOT_LEGAL_CONTENT				= 7		'内容为空或过长oblog.setup(75,0)
Const ERROR_NOT_LEGAL_KEYWORD				= 8		'内容中含有不合法的关键字
Const ERROR_FORBID_UPLOADFILE				= 9		'当前系统设置不允许上传文件
Const ERROR_NOSPACE_FOR_UPLOADFILE			= 10	'上传空间已满，不允许上传文件,请整理上传文档
Const ERROR_NOT_LEGAL_GETRECENTPOSTS_NUM	= 11	'超过允许的获取文章数量
Const ERROR_NOT_LEGAL_PUBLISHSPACE_TIME		= 12	'不符合允许发布的最小时间间隔
Const ERROR_SHUTDOWN_UPLOADFILE				= 13	'不允许上传文件
Const ERROR_SHUTDOWN_UPLOADFILE_1			= 14	'单个文件尺寸超过限制
Const ERROR_SHUTDOWN_UPLOADFILE_2			= 15	'不是合法的上传类型
Const ERROR_LOCKIP							= 16	'用户ip被锁定
Const ERROR_NOT_ADDPOST						= 17	'系统禁止发布日志
Const ERROR_GROUP_ISPOSTMAX					= 18	'用户所在用户组每天发布的日志达到上限

Function ErrorDetail(faultCode)
	select Case faultCode
		Case 1
			ErrorDetail = "不是有效格式的XML请求"
		Case 2
			ErrorDetail = "未知的BlogAPI方法"
		Case 3
			ErrorDetail = "用户名或密码错误"
		Case 4
			ErrorDetail = "要修改的文章不存在"
		Case 5
			ErrorDetail = "ASP端数据库操作失败"
		Case 6
			ErrorDetail = "标题为空或大于100"
		Case 7
			ErrorDetail = "内容为空或过长(不超过" & oblog.CacheConfig(34) & ")"
		Case 8
			ErrorDetail = "内容中含有不合法的关键字"
		Case 9
			ErrorDetail = "当前系统设置不允许上传文件" & afxDebug
		Case 10
			ErrorDetail = "上传空间已满，不允许上传文件,请整理上传文档"
		Case 11
			ErrorDetail = "超过允许获取的最多文章数量"
		Case 12
			ErrorDetail = "不符合允许的两次发布文章的最小时间间隔"
		Case 13
			ErrorDetail = "当前设置不允许上传文件"
		Case 14
			ErrorDetail = "文件尺寸超过限制"
		Case 15
			ErrorDetail = "不是合法的上传类型"
		Case 16
			ErrorDetail = "用户IP被锁定"
		Case 17
			ErrorDetail = "系统临时禁止发布日志"
		Case 18
			ErrorDetail = "超过每日发布日志上限"
		Case Else
			ErrorDetail = "调试代码" & afxDebug
	End select

End Function


Function ResponseError(faultCode)

	Dim strXML
	Dim strError

	strXML="<?xml version=""1.0"" encoding=""gb2312""?><methodResponse><fault><value><struct><member><name>faultCode</name><value><int>$1</int></value></member><member><name>faultString</name><value><string>$2</string></value></member></struct></value></fault></methodResponse>"

	strError=strXML
	strError=Replace(strError,"$1",TransferHTML(faultCode,"[<][>][&][""]"))
	strError=Replace(strError,"$2",TransferHTML(ErrorDetail(faultCode),"[<][>][&][""]"))

	Response.Clear
	Response.Write strError
	Response.End

	conn.Close
	Set conn = Nothing
End Function


Function TransferHTML(source,para)
	On Error Resume Next
	Dim objRegExp

	'先换"&"
	If Instr(para,"[&]")>0 Then  source=Replace(source,"&","&amp;")
	If Instr(para,"[<]")>0 Then  source=Replace(source,"<","&lt;")
	If Instr(para,"[>]")>0 Then  source=Replace(source,">","&gt;")
	If Instr(para,"[""]")>0 Then source=Replace(source,"""","&quot;")
	If Instr(para,"[space]")>0 Then source=Replace(source," ","&nbsp;")
	If Instr(para,"[enter]")>0 Then
		source=Replace(source,vbCrLf,"<br/>")
		source=Replace(source,vbLf,"<br/>")
	End If

	TransferHTML=source

End Function


Function FilterSQL(strSQL)

	FilterSQL=CStr(Replace(strSQL,chr(39),chr(39)&chr(39)))

End Function


Function GetGeneralCategories()

	GetGeneralCategories = False
	Dim i
	Dim aryAllData
	Dim arySingleData()
	Dim rs

	Erase GeneralCategories

	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select [subjectid],[subjectname],[subjectname],[ordernum],[subjectlognum] FROM [oblog_subject] where userid="&objUser.id,conn,1,1
	If (Not rs.bof) And (Not rs.eof) Then
		i=rs.RecordCount
		ReDim GeneralCategories(i)
		aryAllData = rs.GetRows()
		rs.Close
		Set rs = Nothing
		'k = UBound(aryAllData,0)
		'l = UBound(aryAllData,1)
		For i = 0 To i-1
			Set GeneralCategories(i) = New BlogCategory
			GeneralCategories(i).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i)))
		Next
	else
		rs.close
		set rs=nothing
	End If

	GetGeneralCategories = True

End Function


Function GetSystemCategories()

	GetSystemCategories = False
	Dim i
	Dim aryAllData
	Dim arySingleData()
	Dim rs

	Erase SystemCategories

	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select [classid],[classname],[classname],[ordernum],[classlognum] FROM [oblog_logclass] WHERE idType = 0",conn,1,1
	If (Not rs.bof) And (Not rs.eof) Then
		i=rs.RecordCount
		ReDim SystemCategories(i)
		aryAllData = rs.GetRows()
		rs.Close
		Set rs = Nothing
		'k = UBound(aryAllData,0)
		'l = UBound(aryAllData,1)
		For i = 0 To i-1
			Set SystemCategories(i) = New BlogCategory
			SystemCategories(i).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i)))
		Next
	else
		rs.close
		set rs=nothing
	End If

	GetSystemCategories = True

End Function


Sub deloneblog(logid)
	Dim truedel,wsql
	truedel = false
	wsql=" and ( userid="&objUser.Id&" or authorid="&objUser.Id&" )"

    logid = Int(logid)
    Dim uid, delname, rst, fso, sid,Scores
    Set rst = Server.CreateObject("adodb.recordset")
    If Not IsObject(conn) Then link_database
    rst.open "select userid,logfile,subjectid,logtype,scores,isdel from oblog_log where logid="&logid&wsql,conn,1,3
    If rst.Eof Then
        rst.Close
        Set rst = Nothing
        Exit Sub
    End If
	uid = rst(0)
	delname = Trim(rst(1))
	sid = rst(2)
	'清理图片记录,已取消
'	If rst("logtype") = 1 Then
'	    Call DeletePhotos(logid)
'	End If
	'真实域名需要重新整理文件数据
	'物理文件即时删除
	If true_domain = 1 And delname <> "" Then
	    If InStr(delname, "archives") Then
	        delname = Right(delname, Len(delname) - InStrRev(delname, "archives") + 1)
	    Else
	        delname = Right(delname, Len(delname) - InStrRev(delname, "/"))
	    End If
	    delname=oblog.l_udir&"/"&oblog.l_ufolder&"/"&delname
	    'Response.write(delname)
	    'Response.end
	End If
	If delname <> "" Then
	    Set fso = Server.CreateObject(oblog.CacheCompont(1))
	    If fso.FileExists(Server.MapPath(delname)) Then fso.DeleteFile Server.MapPath(delname)
	End If
	Scores=OB_IIF(rst("scores"),0)
	'回收与删除
	'Response.Write(truedel)
	'Response.End()
	If not truedel Then
		rst("isdel")=1
		rst.Update
	Else
		rst.Delete
	End If
	rst.Close
	'--------------------------------------------
	Call Tags_UserDelete(logid)
	'更新计数器
	oblog.Execute ("update oblog_user set log_count=log_count-1 where userid=" & uid)
	If not truedel Then
		oblog.Execute ("Update oblog_comment Set isdel=1 where mainid=" & clng(logid))
	Else
		oblog.Execute ("delete from oblog_comment where mainid=" & clng(logid))
	End If
	oblog.Execute ("update oblog_subject set subjectlognum=subjectlognum-1 where subjectid=" & clng(sid))
	'删除积分
	Call oblog.GiveScore("",-1*Abs(oblog.CacheScores(3)),"")
	'--------------------------------------------
	Dim blog
	Set blog = New class_blog
	blog.userid = uid
	blog.Update_Subject uid
	blog.Update_index 0
	blog.Update_newblog (uid)
	Set blog = Nothing
	Set fso = Nothing
	Set rst = Nothing
End Sub


Class BlogUser

	Public Name
	Public Password
	Public Id
	Public Url

	Public Function Verify()

		Dim strUserName
		Dim strPassWord
		Dim TruePassWord
		TruePassWord = RndPassword(16)

		Verify = False
		strUserName = FilterSQL(Name)
		strPassWord = FilterSQL(Password)		
		oblog.Execute ("UPDATE oblog_user SET TruePassWord = '"&TruePassWord&"' WHERE username = '"&strUserName&"' AND password = '"&strPassWord&"'")
		oblog.SaveCookie strUserName, TruePassWord, 0
		'afxDebug=oblog.checkuserlogined()
		if oblog.checkuserlogined() then
			Id = oblog.l_uid
			Verify = True
		else
			Verify = False
		end if
	End Function

End Class


Class BlogCategory

	Public Id
	Public Name
	Public Intro
	Public Order
	Public Count

	Public Function LoadInfoByArray(aryCateInfo)

		If IsArray(aryCateInfo) = True Then
			Id		= aryCateInfo(0)
			Name	= aryCateInfo(1)
			Intro	= aryCateInfo(2)
			Order	= aryCateInfo(3)
			Count	= aryCateInfo(4)
		End If

		If IsNull(Intro) Then
			Intro=""
		End If

		LoadInfoByArray=True

	End Function

End Class


Class BlogArticle

	Public Id

	Public Topic
	Public Log_Text
	Public Face
	Public AddTime
	Public Tags
	Public Trackback

	Public ClassId
	Public SubjectId
	Public AuthorID
	Public Author
	Public UserId


	Public IsHide
	Public IsTop
	Public TbUrl
	Public LogType
	Public IsEncomment
	Public Abstract
	Public IsPassword
	Public PassCheck
	Public IsDraft
	Public Iis
	Public CommentNum
	Public TrackbackNum
	Public Blog_Password
	Public TrueTime

	Private Function SetDefaultData()

		Topic			= EncodeJP(oblog.filt_astr(Topic,250))
		Log_Text		= EncodeJP(oblog.filtpath(oblog.filt_badword(Log_Text)))
		Face			= 0
		'AddTime			=		'xml传入
		If ClassId = "" Then ClassId = 0 End If		'xml传入
		If SubjectId = "" Then SubjectId = 0 End If	'xml传入
		'AuthorID		=		'由全局变量传入
		'Author			=		'由全局变量传入
		'UserId			=		'由全局变量传入
		IsHide			= 0
		IsTop			= 0
		TbUrl			= ""
		LogType			= 0
		IsEncomment		= 1
		'Abstract		=		'由xml传入
		IsPassword		= ""
		If oblog.l_Group(11,0) = 1 Then'日志需要管理员审核后才可见
			PassCheck = 0
		Else
			PassCheck = 1
		End If
		'IsDraft		= 		'是否为草稿,由xml传入
		Iis				= 0
		CommentNum		= 0
		TrackbackNum	= 0
		Blog_Password	= 0
		TrueTime		= Now()
		'Tags			=		'由xml传入
		'TrackBack		=		'由xml传入

	End Function

	Public Function AddNew()

		AddNew = False
		'系统临时禁止发布日志
		If Application(cache_name_user&"_systemenmod")<>"" Then
			Dim enStr
			enStr=Application(cache_name_user&"_systemenmod")
			enStr=Split(enStr,",")
			If enStr(2)="1" Then ResponseError(ERROR_NOT_ADDPOST):Exit Function
		End If

		SetDefaultData()

		'标题为空或大于100
		If Topic = "" Or StrLength(Topic) > 100 Then
			ResponseError(ERROR_NOT_LEGAL_TITLE)
			Exit Function
		End If
		'内容为空或大于oblog.setup(75,0)
		If Log_Text = "" Or StrLength(Log_Text)>oblog.cacheconfig(34) Then
			ResponseError(ERROR_NOT_LEGAL_CONTENT)
			Exit Function
		End If
		'内容中含有系统不允许发布的关键字
		If oblog.chk_badword(Log_Text) > 0 Then
			ResponseError(ERROR_NOT_LEGAL_KEYWORD)
			Exit Function
		End If

		If StrLength(Tags) > 255 Then'Tags大于255字符置0
			Tags = ""
		End If

		If StrLength(Abstract) > 500 Then'摘要大于500字符置0
			Abstract = ""
		End If

		If (Abstract <> "") Then
			Abstract = oblog.filt_html(Abstract)
		End If

		Dim rs
		Set rs = Server.CreateObject("adodb.recordset")
		rs.open "select TOP 1 * FROM [oblog_log] where Userid="&Userid&" ORDER BY logid desc", conn, 2, 2

		If (Not rs.bof) And (Not rs.eof) Then
			'判断是否超出允许发布的时间间隔
			Dim timeDiff
			'先判断是否超过一天
			timeDiff = DateDiff("d", rs("truetime"), Now())
			If timeDiff = 0 Then'当日内
				timeDiff = DateDiff("s", rs("truetime"), Now())
				If ((timeDiff < MAX_PUBLISHSPACE_TIME) And (MAX_PUBLISHSPACE_TIME <> 0)) Then
					rs.Close
					Set rs = Nothing
					Call ResponseError(ERROR_NOT_LEGAL_PUBLISHSPACE_TIME)
				End If
			End If
		End If

		rs.AddNew

		rs("topic")			= Topic
		rs("logtext")		= Log_Text
		rs("face")			= Face
		rs("addtime")		= AddTime
		rs("classid")		= Classid
		rs("subjectid")		= Subjectid
		rs("authorid")		= AuthorId
		rs("author")		= Author
		rs("userid")		= Userid
		rs("ishide")		= IsHide
		rs("istop")			= IsTop
		rs("tburl")			= TbUrl
		rs("logtype")		= LogType
		rs("isencomment")	= IsEncomment
		rs("abstract")		= Abstract
		rs("ispassword")	= IsPassword
		rs("passcheck")		= PassCheck
		rs("isdraft")		= IsDraft
		rs("iis")			= Iis
		rs("commentnum")	= CommentNum
		rs("trackbacknum")	= TrackbackNum
		rs("blog_password")	= Blog_Password
		rs("truetime")		= TrueTime
		'增加积分
        Call oblog.GiveScore("",oblog.cacheScores(3),"")
        rs("scores")=oblog.cacheScores(3)

		rs.Update
		rs.Close

		Set rs = conn.execute("select max(logid) from oblog_log where userid="&userid)
		Id = rs(0)

		If (Tags <> "") Then
			Call Tags_UserAdd(Tags, Userid, Id)
		End If

		conn.Execute("UPDATE [oblog_user] SET [log_count] = [log_count] + 1 WHERE [userid] = "&AuthorId)

		If (Subjectid > 0) Then
			conn.Execute("UPDATE [oblog_subject] SET [subjectlognum] = [subjectlognum] + 1 WHERE [subjectid] = "&Subjectid)
		End If

		If (Classid > 0) Then
			conn.Execute("UPDATE [oblog_logclass] SET [classlognum] = [classlognum] + 1 WHERE [classid] = "&Classid)
		End If

		conn.Execute("UPDATE [oblog_setup] SET [log_count] = log_count + 1")

		Dim blog
		Set blog = New class_blog

		blog.userid=Userid

		If (TrackBack <> "") Then
			Dim objTrackBack
			Set objTrackBack=New Class_TrackBack
			objTrackBack.Logid=Id
			objTrackBack.Blog_Name=blog.BlogName
			objTrackBack.Title=Topic
			objTrackBack.URL=oblog.setup(3,0)  & "go.asp?logid=" &Id
			objTrackBack.Excerpt=Topic & "<br />oBlog Created"
			Call objTrackBack.ProcessMultiPing(TrackBack)
			Set objTrackBack =Nothing
		End If

		blog.update_log id,0
		blog.update_calendar(id)
		blog.update_newblog(Userid)
		blog.update_subject(Userid)
		blog.update_info(Userid)
		set blog=nothing
		Set rs = Nothing
		AddNew = True

	End Function

	Public Function VerifyId(log_ID)

		VerifyId = False
		'Call CheckParameter(log_ID,"Int",0)

		Dim rs
		Set rs = conn.Execute("select [logid] FROM [oblog_log] WHERE [logid]=" & CLng(log_ID)&" and userid="&objUser.id)

		If (Not rs.bof) And (Not rs.eof) Then
			Id		= rs("logid")
		Else
			VerifyId = False
		End If

		rs.Close
		Set rs = Nothing

		VerifyId = True

	End Function

	Public Function LoadInfobyID(log_ID)

		'Call CheckParameter(log_ID,"Int",0)

		Dim rs
		Set rs = conn.Execute("select * FROM [oblog_log] WHERE [logid]=" & CLng(log_ID)&" and userid="&objUser.id)

		If (Not rs.bof) And (Not rs.eof) Then

			Id			=	rs("logid")
			SubjectId	=	rs("subjectid")
			Topic		=	rs("topic")
			Log_Text	=	rs("logtext")
			AuthorId	=	rs("authorid")
			AddTime		=	rs("addtime")
			CommentNum	=	rs("commentnum")
			TrackbackNum=	rs("trackbacknum")
			Author		=	rs("author")

		Else
			Exit Function
		End If

		rs.Close
		Set rs=Nothing

		LoadInfobyID=True

	End Function

	Public Function Modify()

		Modify = False

		'If ClassId="" then ClassId = 0 End If
		If SubjectId = "" then SubjectId = 0 End If

		Dim rs, oldSubjectId

		Set rs = conn.Execute("select [subjectid] FROM [oblog_log] WHERE [logid] = " & CLng(Id)&" and userid="&objUser.id)
		oldSubjectId = rs("subjectid")

		rs.Close
		Set rs = Nothing

		If (Subjectid > 0) And (Subjectid <> oldSubjectid) Then
			conn.Execute("UPDATE [oblog_subject] SET [subjectlognum] = subjectlognum + 1 WHERE [subjectid] = "&Subjectid)
			conn.Execute("UPDATE [oblog_subject] SET [subjectlognum] = subjectlognum - 1 WHERE [subjectid] = "&oldSubjectid)
		End If

		conn.Execute("UPDATE [oblog_log] SET [topic]='"&Topic&"',[logtext]='"&Log_Text&"',[truetime]='"&Now()&"',[subjectid]="&SubjectId&" WHERE [logid] =" & Id &" and userid="&objUser.id)

		Modify = True

	End Function

	Public Function Delete()

		deloneblog(Id)
		Delete = True

	End Function

	Public Function StrLength(Str)

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
            StrLength = t
        Else
        StrLength = Len(Str)
        End If
    End Function

End Class


Class BlogUpLoadFile

	Public ID
	Public AuthorID

	Public FileSize
	Public FileName
	Public PostTime
	Public Stream

	Public BackUrl

	Private FUploadType
	Public Property Let UploadType(strUploadType)
		If (strUploadType="Stream") Then
			FUploadType=strUploadType
		Else
			FUploadType="Form"
		End If
	End Property
	Public Property Get UploadType
		If IsEmpty(FUploadType)=True Then
			UploadType="Form"
		Else
			UploadType = FUploadType
		End If
	End Property


	Private Function UpLoad_Stream()

		FileSize=LenB(Stream)

	End Function


	Public Function UpLoad()
		UpLoad = FALSE
		Dim enupload, upfiletype, onesize, maxsize, freesize, FileExt, upload_dir
		upload_dir=oblog.CacheConfig(56)
		If UploadType="Form" Then
			'Call UpLoad_Form()
		ElseIf UploadType="Stream" Then
			Call UpLoad_Stream()
		End If

		If oblog.l_Group(24,0)=-1 Then
			enupload=0
		Else
			enupload=1
		End If
		upfiletype=oblog.l_Group(22,0)
		'Response.Write upfiletype
		'Response.End
		onesize=oblog.l_Group(23,0)
		maxsize=oblog.l_Group(24,0)
		'当前系统设置不允许上传文件
		If enupload = 0 Then
			ResponseError(ERROR_FORBID_UPLOADFILE)
			Exit Function
		End If
		'上传空间已满，不允许上传文件,请整理上传文档
		'If  objUser.UserUpMax > 0 Then maxsize = objUser.UserUP
		freesize = Int(maxsize - oblog.l_uUpUsed / 1024)
		If freesize <= 0 Then
			ResponseError(ERROR_NOSPACE_FOR_UPLOADFILE)
			Exit Function
		End if

		FileName=FilterSQL(FileName)

		Dim filePath
		If upload_dir <> "" Then
			filePath = upload_dir
		Else
			filePath = oblog.l_udir&"/"&oblog.l_ufolder&"/upload"
		End If
		filePath = CreatePath(filePath)

		Randomize
		FileExt=Lcase(Mid(FileName,InStrRev(FileName, ".")+1))
		if not CheckFileExt(FileExt,upfiletype) then
			ResponseError(ERROR_NOSPACE_FOR_UPLOADFILE)
			Exit Function
		end if
		FileExt=FixName(FileExt)
		FileName = Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2) & Right("0"&Second(Now),2) & Int(9 * Rnd) & Int(9 * Rnd) & Int(9 * Rnd) & Int(9 * Rnd) &"."& FileExt

		Dim objStreamFile
		Set objStreamFile = Server.CreateObject(oblog.CacheCompont(2))

		objStreamFile.Type = 1
		objStreamFile.Mode = 3
		objStreamFile.Open
		objStreamFile.Write Stream
		if objStreamFile.size/1024>freesize then
			objStreamFile.Close
			ResponseError(ERROR_NOSPACE_FOR_UPLOADFILE)
			Exit Function
		elseif objStreamFile.size/1024>onesize then
			objStreamFile.Close
			ResponseError(ERROR_SHUTDOWN_UPLOADFILE_1)
			Exit Function
		else
			objStreamFile.SaveToFile Server.MapPath(filePath) & "\" & FileName,2
		end if
		objStreamFile.Close

		Dim rs
		'If instr("jpg,gif,bmp,pcx,png,psd",fileExt) = 0 Then isphoto = 0

		conn.execute("UPDATE [oblog_user] SET [user_upfiles_size] = user_upfiles_size + "&FileSize&" WHERE [userid] = "&AuthorId)
		Set rs = Server.CreateObject("adodb.recordset")
		rs.open "select TOP 1 * FROM [oblog_upfile]", conn, 2, 2

		rs.AddNew

		rs("userid")		= AuthorId
		rs("file_name")		= FileName
		rs("file_path")		= Mid(filePath, 4) & FileName
		rs("file_ext")		= fileExt
		rs("file_size")		= FileSize
		rs("isphoto")		= 0
		rs.Update

		rs.Update
		rs.close
		Set rs = Nothing

		BackUrl = oblog.cacheconfig(3) & Mid(filePath, 4) & FileName

		UpLoad = True

	End Function

	'检查上传目录，若无目录则自动建立
	Public Function CreatePath(PathValue)

		Dim objFSO,Fsofolder,uploadpath,upload_dir
		upload_dir=oblog.CacheConfig(56)
		If upload_dir<>"" Then
			uploadpath = year(Date) & "-" & month(Date)
		Else
			uploadpath=""
		End If
		If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"

		Set objFSO = Server.CreateObject(oblog.CacheCompont(1))


		If objFSO.FolderExists(Server.MapPath("../" & PathValue & uploadpath))=False Then
			objFSO.CreateFolder Server.MapPath("../" & PathValue & uploadpath)
		End If
		If upload_dir<>"" Then
			CreatePath = "../" & PathValue & uploadpath & "/"
		Else
			CreatePath = "../" & PathValue
		End If
		Set objFSO = Nothing

	End Function

	Private Function CheckFileExt(FileExt,upfiletype)
		Dim Forumupload,i
		CheckFileExt=False
		If FileExt="" or IsEmpty(FileExt) Then
			CheckFileExt = False
			Exit Function
		End If
		select Case LCase(FileExt)
				Case "asp","asa","aspx","shtm","shtml","php","php3","jsp"
					CheckFileExt = False
					Exit Function
				Case else
		End select
		Forumupload = Split(upfiletype,"|")
		For i = 0 To ubound(Forumupload)
			If FileExt = Trim(Forumupload(i)) Then
				CheckFileExt = True
				Exit Function
			Else
				CheckFileExt = False
			End If
		Next
	End Function
	Private Function FixName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		FixName = Lcase(UpFileExt)
		FixName = Replace(FixName,Chr(0),"")
		FixName = Replace(FixName,".","")
		FixName = Replace(FixName,"'","")
		FixName = Replace(FixName,"asp","_")
		FixName = Replace(FixName,"asa","_")
		FixName = Replace(FixName,"aspx","_")
		FixName = Replace(FixName,"cer","_")
		FixName = Replace(FixName,"cdx","_")
		FixName = Replace(FixName,"htr","_")
		FixName = Replace(FixName,"shtm","_")
		FixName = Replace(FixName,"shtml","_")
		FixName = Replace(FixName,"php","_")
		FixName = Replace(FixName,"php3","_")
		FixName = Replace(FixName,"jsp","_")
	End Function
End Class
%>
