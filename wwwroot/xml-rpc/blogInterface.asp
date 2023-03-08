<!--#include file="../inc/inc_syssite.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<!--#include file="../inc/class_Trackback.asp"-->
<%
Dim afxDebug
Const MAX_GETRECENTPOSTS_NUM				= 0		'getRecentPosts����������������,0Ϊ������
Const MAX_PUBLISHSPACE_TIME					= 10		'���η���������Сʱ���� ��λ/��, 0Ϊ������
'��ЧConst MAX_UPLOADFILESPACE_TIME				= 0		'�����ϴ��ļ���Сʱ���� ��λ/��, 0Ϊ������

Const UPLOADFILE_SIGN						= true	'�Ƿ������ϴ��ļ�.

Const ERROR_NOT_LEGAL_XMLREQUEST			= 1		'������Ч��ʽ��XML����
Const ERROR_UNKNOW_BLOGAPIMETHOD			= 2		'δ֪��BlogAPI����
Const ERROR_NOT_LEGAL_USER					= 3		'�û������������
Const ERROR_NOT_EXIST_ARTICLE				= 4		'Ҫ�޸ĵ����²�����
Const ERROR_ACCESS_DATABASE_FAILED			= 5		'ASP�����ݿ����ʧ��
Const ERROR_NOT_LEGAL_TITLE					= 6		'����Ϊ�ջ����100
Const ERROR_NOT_LEGAL_CONTENT				= 7		'����Ϊ�ջ����oblog.setup(75,0)
Const ERROR_NOT_LEGAL_KEYWORD				= 8		'�����к��в��Ϸ��Ĺؼ���
Const ERROR_FORBID_UPLOADFILE				= 9		'��ǰϵͳ���ò������ϴ��ļ�
Const ERROR_NOSPACE_FOR_UPLOADFILE			= 10	'�ϴ��ռ��������������ϴ��ļ�,�������ϴ��ĵ�
Const ERROR_NOT_LEGAL_GETRECENTPOSTS_NUM	= 11	'��������Ļ�ȡ��������
Const ERROR_NOT_LEGAL_PUBLISHSPACE_TIME		= 12	'����������������Сʱ����
Const ERROR_SHUTDOWN_UPLOADFILE				= 13	'�������ϴ��ļ�
Const ERROR_SHUTDOWN_UPLOADFILE_1			= 14	'�����ļ��ߴ糬������
Const ERROR_SHUTDOWN_UPLOADFILE_2			= 15	'���ǺϷ����ϴ�����
Const ERROR_LOCKIP							= 16	'�û�ip������
Const ERROR_NOT_ADDPOST						= 17	'ϵͳ��ֹ������־
Const ERROR_GROUP_ISPOSTMAX					= 18	'�û������û���ÿ�췢������־�ﵽ����

Function ErrorDetail(faultCode)
	select Case faultCode
		Case 1
			ErrorDetail = "������Ч��ʽ��XML����"
		Case 2
			ErrorDetail = "δ֪��BlogAPI����"
		Case 3
			ErrorDetail = "�û������������"
		Case 4
			ErrorDetail = "Ҫ�޸ĵ����²�����"
		Case 5
			ErrorDetail = "ASP�����ݿ����ʧ��"
		Case 6
			ErrorDetail = "����Ϊ�ջ����100"
		Case 7
			ErrorDetail = "����Ϊ�ջ����(������" & oblog.CacheConfig(34) & ")"
		Case 8
			ErrorDetail = "�����к��в��Ϸ��Ĺؼ���"
		Case 9
			ErrorDetail = "��ǰϵͳ���ò������ϴ��ļ�" & afxDebug
		Case 10
			ErrorDetail = "�ϴ��ռ��������������ϴ��ļ�,�������ϴ��ĵ�"
		Case 11
			ErrorDetail = "���������ȡ�������������"
		Case 12
			ErrorDetail = "��������������η������µ���Сʱ����"
		Case 13
			ErrorDetail = "��ǰ���ò������ϴ��ļ�"
		Case 14
			ErrorDetail = "�ļ��ߴ糬������"
		Case 15
			ErrorDetail = "���ǺϷ����ϴ�����"
		Case 16
			ErrorDetail = "�û�IP������"
		Case 17
			ErrorDetail = "ϵͳ��ʱ��ֹ������־"
		Case 18
			ErrorDetail = "����ÿ�շ�����־����"
		Case Else
			ErrorDetail = "���Դ���" & afxDebug
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

	'�Ȼ�"&"
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
	'����ͼƬ��¼,��ȡ��
'	If rst("logtype") = 1 Then
'	    Call DeletePhotos(logid)
'	End If
	'��ʵ������Ҫ���������ļ�����
	'�����ļ���ʱɾ��
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
	'������ɾ��
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
	'���¼�����
	oblog.Execute ("update oblog_user set log_count=log_count-1 where userid=" & uid)
	If not truedel Then
		oblog.Execute ("Update oblog_comment Set isdel=1 where mainid=" & clng(logid))
	Else
		oblog.Execute ("delete from oblog_comment where mainid=" & clng(logid))
	End If
	oblog.Execute ("update oblog_subject set subjectlognum=subjectlognum-1 where subjectid=" & clng(sid))
	'ɾ������
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
		'AddTime			=		'xml����
		If ClassId = "" Then ClassId = 0 End If		'xml����
		If SubjectId = "" Then SubjectId = 0 End If	'xml����
		'AuthorID		=		'��ȫ�ֱ�������
		'Author			=		'��ȫ�ֱ�������
		'UserId			=		'��ȫ�ֱ�������
		IsHide			= 0
		IsTop			= 0
		TbUrl			= ""
		LogType			= 0
		IsEncomment		= 1
		'Abstract		=		'��xml����
		IsPassword		= ""
		If oblog.l_Group(11,0) = 1 Then'��־��Ҫ����Ա��˺�ſɼ�
			PassCheck = 0
		Else
			PassCheck = 1
		End If
		'IsDraft		= 		'�Ƿ�Ϊ�ݸ�,��xml����
		Iis				= 0
		CommentNum		= 0
		TrackbackNum	= 0
		Blog_Password	= 0
		TrueTime		= Now()
		'Tags			=		'��xml����
		'TrackBack		=		'��xml����

	End Function

	Public Function AddNew()

		AddNew = False
		'ϵͳ��ʱ��ֹ������־
		If Application(cache_name_user&"_systemenmod")<>"" Then
			Dim enStr
			enStr=Application(cache_name_user&"_systemenmod")
			enStr=Split(enStr,",")
			If enStr(2)="1" Then ResponseError(ERROR_NOT_ADDPOST):Exit Function
		End If

		SetDefaultData()

		'����Ϊ�ջ����100
		If Topic = "" Or StrLength(Topic) > 100 Then
			ResponseError(ERROR_NOT_LEGAL_TITLE)
			Exit Function
		End If
		'����Ϊ�ջ����oblog.setup(75,0)
		If Log_Text = "" Or StrLength(Log_Text)>oblog.cacheconfig(34) Then
			ResponseError(ERROR_NOT_LEGAL_CONTENT)
			Exit Function
		End If
		'�����к���ϵͳ���������Ĺؼ���
		If oblog.chk_badword(Log_Text) > 0 Then
			ResponseError(ERROR_NOT_LEGAL_KEYWORD)
			Exit Function
		End If

		If StrLength(Tags) > 255 Then'Tags����255�ַ���0
			Tags = ""
		End If

		If StrLength(Abstract) > 500 Then'ժҪ����500�ַ���0
			Abstract = ""
		End If

		If (Abstract <> "") Then
			Abstract = oblog.filt_html(Abstract)
		End If

		Dim rs
		Set rs = Server.CreateObject("adodb.recordset")
		rs.open "select TOP 1 * FROM [oblog_log] where Userid="&Userid&" ORDER BY logid desc", conn, 2, 2

		If (Not rs.bof) And (Not rs.eof) Then
			'�ж��Ƿ񳬳���������ʱ����
			Dim timeDiff
			'���ж��Ƿ񳬹�һ��
			timeDiff = DateDiff("d", rs("truetime"), Now())
			If timeDiff = 0 Then'������
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
		'���ӻ���
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
        WINNT_CHINESE = (Len("�й�") = 2)
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
		'��ǰϵͳ���ò������ϴ��ļ�
		If enupload = 0 Then
			ResponseError(ERROR_FORBID_UPLOADFILE)
			Exit Function
		End If
		'�ϴ��ռ��������������ϴ��ļ�,�������ϴ��ĵ�
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

	'����ϴ�Ŀ¼������Ŀ¼���Զ�����
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
