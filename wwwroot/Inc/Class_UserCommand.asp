<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="../conn.asp"-->
<!--#include  file="class_sys.asp"-->
<!--#include  file="Inc_Calendar.asp"-->
<!--#include  file="Inc_ubb.asp"-->
<%
Dim oBlog
Set oBlog = New class_sys
oBlog.start

'�û�ȫ������ģ��
'��ģ��Ŀǰ������������ģ�����
Class Class_UserCommand
	Public Action
	Public ID,FileID
	Public rst
	Public Title
	Public ErrMsg
	Public mUserSkinLog,mYear,mMonth,mDay
	Private mUserName,mUserId,mUserPath,mUserNickName,mUserFolder,mBlogName,mUserPhotoRow,mUsersublist,mUserCmdpath,mUserLogPath,mUserIndexlist,mUserIcon1
	Private strLogN,strUrl,ShowDigg
	Private Sql,SqlStart,SqlPart,SqlEnd,rstSubject,strErrMsg,strPlayerUrl

	Private Sub Class_Initialize()
		userid=Clng(Request("uid"))
		strPlayerUrl= blogurl & "PhotoPlayer.asp?userid="&mUserid
		'G_P_PerMax=5
	End Sub

	Private Sub Class_Terminate()

	End Sub

	Public Property Let userid(ByVal Values)
		Dim rstmp, strSql
		mUserid = CLng(Values)
		'SqlStart = "select  * From oblog_log Where userid="& mUserId & " "
		SqlStart = "select  * From oblog_log Where 1=1 "
		'SqlEnd="  And ishide=0 and passcheck=1 and isdraft=0 and blog_password=0 Order by istop,addtime Desc"
		SqlEnd=" and passcheck=1 and isdraft=0 and isdel=0  And (userid="& mUserId & " Or authorid=" & mUserId & ") Order by istop Desc,addtime Desc"
		Action=LCase(Request("do"))
		Id=OB_IIF(Request("Id"),0)
		Call GetUserInfo()
		G_P_FileName=mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do="

		Dim rsUser
		Set rsUser = oblog.Execute ("SELECT isdigg FROM oblog_user WHERE userid = "&OB_IIF(mUserid,0))
		If Not rsUser.Eof Then
			If OB_IIF (rsUser(0),1) = 1 Then
				ShowDigg = vbcrlf & "<div class=""digg_list"" style=""float: right; display:inline; margin: 0 10px 5px 0; width: 45px; height: 55px; background: url("&blogurl&"Images/digg.gif) no-repeat left top; text-align: center; "">" & vbcrlf
				ShowDigg = ShowDigg & "	<div class=""digg_number"" style=""width:45px;padding: 10px 0 11px 0;font-size:18px;font-weight:600;color:#333;font-family:tahoma,Arial,Helvetica,sans-serif;line-height:1.0;"">$diggnum$</div>" & vbcrlf
				ShowDigg = ShowDigg & "	<div class=""digg_submit"" style="" padding: 3px 0 0 6px;line-height:1.0;letter-spacing: 6px; ""><a href=""javascript:void(null)"" onclick=""diggit($logid$);"" style=""font-size:12px;line-height:1.0;"">$showmsg$</a></div>" & vbcrlf
				ShowDigg = ShowDigg & "</div>" & vbcrlf
			End if
		Else
			ShowDigg = ""
		End if
	End Property

	Private Function ShowErrorMsg(ByVal strMsg)
		Response.Write oblog.htm2js_div(filtskinpath(strMsg),"oblog_usercontent")
		Response.End
	End Function

	Public Function Process()
		Dim strReturn,strMonth,strDay
		Id=CheckInt(Id)
		strMonth=Request.QueryString("month")
		strDay=Request.QueryString("day")
		'Response.Write "����2��" & Action & "<BR/>" & vbCrlf
		'Response.Write "���2��" & Id & "<BR/>" & vbCrlf
		select Case Action
			Case "index"
				SqlPart=" "
				Sql=SqlStart &	SqlEnd
				G_P_FileName = G_P_FileName & "index"
				strReturn = ShowList(Sql,"ƪ��־","0")
			Case "blogs"
				If Id="" OR Id=0 Then
					SqlPart=" And logType=0"
					G_P_FileName = G_P_FileName & "blogs"
				Else
					SqlPart=" And logType=0 And Subjectid=" & Id
					G_P_FileName = G_P_FileName & "blogs&id=" & Id
				End If
				'SqlPart = SqlPart &" AND (isspecial = 0  OR isspecial IS NULL )"
				Sql=SqlStart & SqlPart & SqlEnd
				strReturn = ShowList(Sql,"ƪ��־","0")
			Case "month"
				Dim LastDay
				G_P_FileName = G_P_FileName & "month&month=" & strMonth
				If Len(strMonth)<>6 OR IsNumeric(strMonth)=False Then
					ErrMsg = "<center>������·����ݣ�ӦΪYYYYMM��ʽ���磺200508</center>"
					ShowErrorMsg ErrMsg
				End If
				strDay=CLng(Left(strMonth,4)) & "-" & CLng(Right(strMonth,2)) & "-01"
				mYear=CLng(Left(strMonth,4))
				mMonth=CLng(Right(strMonth,2))
				If InStr ("01,03,05,07,08,10,12",mMonth)> 0 Then
					LastDay = "31"
				Else
					If mMonth <> "02" Then
						LastDay = "30"
					Else
						If mYear Mod 4 = 0 Then
							LastDay = "29"
						Else
							LastDay = "28"
						End if
					End if
				End if
				If Not IsDate(strDay) Then
					ErrMsg = "<center>������������ݣ�ӦΪYYYYMMDD��ʽ���磺2005-08-01</center>"
					ShowErrorMsg ErrMsg
				End If
				If Is_Sqldata = 0 Then
					SqlPart = " And Datediff("&G_Sql_m&",Addtime,'" & strDay &"')=0"
				Else
					SqlPart = " And Addtime >='"&strMonth&"01' AND Addtime < '"&strMonth&LastDay&"' "
				End if
				Sql=SqlStart & SqlPart & SqlEnd
				strReturn = ShowList(Sql,"ƪ��־","0")
			Case "day"
				G_P_FileName = G_P_FileName & "day&day=" & strDay
				mYear=CLng(Year(strDay))
				mMonth=CLng(Month(strDay))
				If Not IsDate(strDay) Then
					ErrMsg = "<center>��������ڸ�ʽ��ӦΪYYYYMMDD��ʽ���磺2005-08-01</center>"
					ShowErrorMsg ErrMsg
				End If
				If Is_Sqldata = 0 Then
					SqlPart = "And Datediff("&G_Sql_d&",Addtime,'" & strDay &"')=0"
				Else
					SqlPart = "AND Addtime >= '"&GetDateCode(strDay,0)&"' AND Addtime <'"&GetDateCode(CDate(strDay)+1,0)&"' "
				End if
				Sql=SqlStart & SqlPart & SqlEnd
				strReturn = ShowList(Sql,"ƪ��־","0")
			Case "message"
				Sql="select * from oblog_message where userid=" & mUserId & " order by messageid desc"
				G_P_FileName = G_P_FileName & "message"
				strReturn = ShowList(Sql,"������","1")
			Case "comment"
			Case "tag_blogs" '�˴�����־�����ϲ���ʾ
				G_P_FileName = G_P_FileName & "tag_blogs&id=" & Id
				Sql="select a.userid,b.* From "
				Sql=Sql & " (select logid,userid From oblog_usertags Where userid=" & mUserId  & " and tagid=" & id &") a ,"
				'Sql=Sql & " (select * From oblog_log where userid=" & mUserId & " And logType=0) b Where a.logid=b.logid "
				Sql=Sql & " (select * From oblog_log where userid=" & mUserId & ") b Where a.logid=b.logid "
				Sql=Sql & " order By b.addtime Desc"
				strReturn = ShowList(Sql,"ƪ��־","0")
			Case "tag_photos"
				G_P_FileName = G_P_FileName & "tag_photos&id=" & Id
				Sql="select a.userid,b.* From "
				Sql=Sql & " (select logid,userid From oblog_usertags Where userid=" & mUserId  & " and tagid=" & id &") a ,"
				Sql=Sql & " (select * From oblog_log where userid=" & mUserId & " And logType=1) b Where a.logid=b.logid "
				Sql=Sql & " order By b.addtime Desc"
				strReturn = ShowList(Sql,"ƪ��־","0")
			Case "tags"
				strReturn = GetUserTags()
			Case "show"
				strReturn = ShowOneBlog(Id,0)
			Case "album"
				If oblog.CacheConfig(76) = "0" Then
					ErrMsg = "�˹����ѱ�ϵͳ�رգ�"
					ShowErrorMsg ErrMsg
				End if
				G_P_FileName = G_P_FileName & "album&id=" &Id
				if id>0 then
					Sql = "select photo_path,fileID,photo_Title,photo_name from oblog_album where TeamID=0 and (ishide=0 OR ishide IS NULL) and userid="&mUserId&" and userClassId="&id&"  order by photoID desc"
				Else
					If id = -1 Then
					'��ʾ���з����ص���Ƭ
						Sql = "select photo_path,fileID,photo_Title,photo_name from oblog_album where TeamID=0 and (ishide=0 OR ishide IS NULL) and userid="&mUserId&" and userClassId=0  order by photoID desc"
					ElseIf id = -2 Then
						Sql = "select photo_path,fileID,photo_Title,photo_name from oblog_album where TeamID=0 and (ishide=0 OR ishide IS NULL) and userid="&mUserId&" order by photoID desc"
					Else
					'��ʾ���з����ص����
						Sql = "SELECT c.photo_path,c.subjectid,c.subjectlognum FROM "
						Sql = Sql &" oblog_subject AS c "
						Sql = Sql &" WHERE c.subjecttype = 1 AND ( c.ishide = 0  OR c.ishide IS NULL) AND c.userid="&mUserId
						Sql = Sql &" ORDER BY c.subjectid DESC"
					End if
				end if
				strReturn = ShowList(sql,"����Ƭ","2")
			Case "flash"
				If oblog.CacheConfig(76) = "0" Then
					ErrMsg = "�˹����ѱ�ϵͳ�رգ�"
					ShowErrorMsg ErrMsg
				End if
				'by Ҷ��
	'			strReturn="��<a href=""#"" onclick=""window.open('"&strPlayerUrl&"','_photo','height=500, width=480, top=100, left=400, toolbar=no, menubar=no, scrollbars=no, resizable=yes,status=no')"">�����Զ�����</a>" & VBCRLF
				strReturn=strReturn&"  <a href='"&mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=album'>��᷽ʽ���</a>"
				strReturn = strReturn &"<div style=""margin:0;width:500px;text-align:center;""><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0' width='100%' height='500' align='middle'><param name=""wmode"" value=""transparent"" /><param name='movie' value='"&blogurl&"photo.swf?blogurl="&blogurl&"&userid="&mUserId&"&f_ext="&f_ext&"' /><param name='quality' value='high' /><embed src='"&blogurl&"photo.swf?blogurl="&blogurl&"&userid="&mUserId&"&f_ext="&f_ext&"' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='100%' height='500'></embed></object></div>"
				'strReturn = strReturn &"<br/>	<div id=""PlayerContainer"" style=""position:absolute;background-color:#fff;z-index:1000;width:600px;height:480px;padding:0px;"" align=""center""><object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0"" width=""100%"" height=""100%"" id=""photoview"" align=""middle""><param name=""allowScriptAccess"" value=""always"" />	<param name=""movie"" value="""&blogurl&"PhotoViewer.swf?blogurl="&blogurl&"&userid=1"" /><param name=""quality"" value=""high"" />	<param name=""wmode"" value=""transparent"" />	<param name=""bgcolor"" value=""#ffffff"" />	<embed src="""&blogurl&"PhotoViewer.swf?blogurl="&blogurl&"&userid=1"" quality=""high"" wmode=""transparent"" bgcolor=""#ffffff"" width=""100%"" height=""100%"" name=""photoview"" align=""middle"" allowScriptAccess=""always"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" /></object>	</div>" & vbCrlf
			Case "info"
				strReturn=showinfo()
			'��ȡ��������б�
			Case "photocomment"
				If oblog.CacheConfig(76) = "0" Then
					ErrMsg = "�˹����ѱ�ϵͳ�رգ�"
					ShowErrorMsg ErrMsg
				End if
				G_P_FileName = G_P_FileName & "photocomment&fileid=" & FileID
				Sql ="select comment_user,homepage,commenttopic,comment,commentid,a.addtime,isguest "
				Sql = Sql & "FROM oblog_albumcomment a,oblog_album b "
				Sql = Sql & "WHERE a.mainid = b.fileid AND iState=1 AND mainid="&FileID
				Sql = Sql & " ORDER By A.addtime DESC"
				strReturn = strReturn & ShowList(Sql,"ƪ����","3")
			Case Else
				SqlPart=" "
				Sql=SqlStart &	SqlEnd
				strReturn = ShowList(Sql,"ƪ��־","0")
		End select
		strReturn=oblog.htm2js_div(filtskinpath(strReturn),"oblog_usercontent")
		Process=strReturn
		'Process="document.write('" & strReturn & "');"
	End Function

	Public Function CreateCalendar()
		Dim strReturn
		If mYear="" Then
			mYear=Year(Date)
			mMonth=Month(Date)
		End If
		strReturn=oblog.htm2js_div(Calendar(mYear,mMonth,mUserId),"calendar")
		CreateCalendar=strReturn
	End Function

	Private Function ShowUserBlogs(rst)
		Dim strBlogs
		Do While Not rst.Eof
			strBlogs= strBlogs & GetOneBlogInfo(rst,"")	& "<BR/>"
			rst.Movenext
		Loop
		'����ͳһ��������ÿƪ����
		strLogMore=Replace(strLogMore,"$show_blogtag$","")
		strLogMore=Replace(strLogMore,"$show_blogzhai$","")
		strLogMore=Replace(strLogMore,"$show_blogtag","")
		strLogMore=filt_inc(strLogMore)
		strLogMore=strLogMore & "<script src="""&BlogDir&"count.asp?action=logs&id="&strLogN&"""></script>"
		'strLogMore=Replace(user_skin_main,"$show_log$",strLogMore)
		ShowUserBlogs= strLogMore
		strLogMore=""
		'��ҳ������ÿ��ֻȡ����������G_P_PerMax������ȫ��ȡ��
	End Function

	Private Function ShowOneBlog(BlogId,isPower)
		Set rst=oblog.Execute("select * From oBlog_log Where logid=" & BlogId)
		ShowOneBlog=GetOneBlogInfo(rst,"1")
	End FUnction

	'��ȡһƪ��־����������
	'ע��ժҪ/�����Լ�β����ǩ�Ĵ���
	Public Function GetOneBlogInfo(byref rst,byval strMode)
		Dim strTopic,strEmot,strAddtime,strLogtext,strAuthor,strLogInfo,strMore
		Dim strOneLog,strTopictxt,strLogMore,show,rssubject,strTmp,xmlstr,rstmp,strart,i
		'����
		'If rst("face")="0" Then strEmot="" Else	strEmot="<img src="&blogurl&"images/face/" & rst("face") &".gIf />"
		'����
		If mUserNickName=""  Then
			strAuthor=mUserName
		Else
			strAuthor=mUserNickName
		End If
		If rst("authorid")<>mUserId Then
			If Not IsNull(rst("author")) Then
				strAuthor=rst("author")
			End If
		End If
		strAddtime=rst("addtime")
		strTopic=strEmot
		If rst("istop")=1 Then strTopic="[�ö�]"
		If rst("subjectid")>0 Then
			rstSubject.Filter="subjectid=" & rst("subjectid")
			If Not rstSubject.Eof Then
				strTopic=strTopic & "<a href="""& BlogDir & UserPath &"/cmd."&f_ext&"?do=subject&id="">["&oblog.filt_html(rssubject(1))&"]</a>"
			End If
		End If

		Dim digg
		digg = ShowDigg
		digg = Replace(digg,"$diggnum$",OB_IIF(rst("DIGGNum"),0))
		digg = Replace(digg,"$logid$",rst("logid"))
		digg = Replace(digg,"$showmsg$","�Ƽ�")

		strTopictxt="<a href="""& BlogDir & rst("logfile")& """>" & oblog.filt_html(rst("topic")) & "</a>"
		If rst("isbest")=1 Then strTopictxt = strTopictxt & "��<img src=../../images/jhinfo.gIf >"
		strTopic = strTopic & strTopictxt
		If rst("istop")=1 Then strTopictxt = "[�ö�]" & strTopictxt
		strLogInfo = strAuthor & " ������ " & strAddtime
		strMore = "<a href="""& BlogDir & rst("logfile")&""">�Ķ�ȫ��<span id=""ob_logr" & rst("logid") & """></span></a>"
		strMore = strMore&" | "&"<a href=""" & BlogDir & rst("logfile")&"#comment"">�ظ�<span id=""ob_logc" & rst("logid") & """></span></a> | <a href=""javascript:void(null);"" onclick=""ListMenu('"&rst("logid")&"')""><span id=""a_"&rst("logid")&""">��ӳ����</span></a><span id=""menu_"&rst("logid")&"""></span>"
		strMore = strMore&" | "&"<a href=""../../showtb.asp?id=" & rst("logid") & """ target=""_blank"">����ͨ��<span id=""ob_logt" & rst("logid") & """></span></a>"
		'ժҪ
		'If Not IsNull(rst("Abstract")) Then
		'	strLogtext=rst("Abstract")
		'Else
			strLogtext="<span id=""ob_logd"& rst("logid") &""">"&digg&"</span>"&rst("logtext")
		'End If
		'�������м����ۼ�
		strLogN=strLogN&"$"&rst("logid")

		'��������ģ������
		strOneLog = Replace(mUserSkinLog,"$show_topic$",strTopic)
		strOneLog = Replace(strOneLog,"$show_loginfo$",strLogInfo)
		strOneLog = Replace(strOneLog,"$show_logtext$",strLogtext)
		strOneLog = Replace(strOneLog,"$show_more$",strMore)
		strOneLog = Replace(strOneLog,"$show_emot$",strEmot)
		'strOneLog = Replace(strOneLog,"$show_author$",strAuthor)
		strOneLog = Replace(strOneLog,"$show_addtime$",strAddtime)
		strOneLog = Replace(strOneLog,"$show_topictxt$",strTopictxt)
		strLogMore=strLogMore&strOneLog
		If strMode="1" Then
			strLogMore=Replace(strLogMore,"$show_blogtag$","")
			strLogMore=Replace(strLogMore,"$show_blogzhai$","")
			strLogMore=Replace(strLogMore,"$show_blogtag","")
			'strLogMore=filt_inc(strLogMore)
			strLogMore=strLogMore & "<script src="""&BlogDir&"count.asp?action=logs&id="&strLogN&"""></script>"
		End If
		GetOneBlogInfo = strLogMore
	End Function

	'�û�TAG�������з�ҳ(Cloud),���ݱ�ǩ��ѯ�������ݲ�������־�������
	Private Function GetUserTags()
		Dim sContent,sSql,rst,iFont,iFontSize
		sSql = "select a.TagId,a.Name,b.TagNum From oblog_tags a,"
		sSql = sSql & "(select Count(*) as TagNum,TagId From oblog_UserTags Where userid=" & mUserId & " Group By TagId ) b Where "
		sSql = sSql & "a.tagid=b.tagid "
		'Response.Write sSql
		Set rst=conn.Execute(sSql)
		If rst.Eof Then
			sContent=""
		Else
			Do While Not rst.Eof
				'����Ϊ10
				iFont=rst("TagNum") Mod 10
				If iFont=0 Then iFontSize=9
				If iFont>-1 And iFont<40 Then iFontSize=12 + iFont
				If iFont >40 Then iFontSize=42
				sContent= sContent & "<li><span><a href="""&mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=tag_blogs&id=" & rst("tagID") & """><font style=""font-size:"& iFontSize &"px;"">" & rst("Name")& "</font></a></span><br />"
				sContent= sContent & "<a href="&blogurl&"tags.asp?tagid=" & rst("tagID") &" target=_blank><img src="&blogurl&"images/icon_blogs.gif border=0 title='��վʹ�ù��ñ�ǩ����־'/></a>"
				sContent= sContent & "<a href="&blogurl&"tags.asp?t=user&tagid=" & rst("tagID") &" target=_blank><img src="&blogurl&"images/icon_users.gif border=0 title='��վʹ�ù��ñ�ǩ���û�'/></a></li>"
				rst.Movenext
			Loop
		End If
		rst.Close
		Set rst=Nothing
		GetUserTags="<div id=""ob_usertags""><ul>"&sContent&"</ul></div>"
		sContent=""
	End Function

	Private Function ShowList(strSql,strUnit,strMode)
		Dim strReturn
		if action="photos" or action="album" then strReturn="<div id=""albumtop""><ul>"&GetUserClasses(action)&"<ul></div>"
		If Request("page")<>"" Then
			G_P_This=CLng(Request("page"))
		Else
			G_P_This=1
		End If
		If strMode = 4 Then ID = -1
		If Not IsObject(conn) Then link_database
		Set rst=Server.CreateObject("Adodb.RecordSet")
		rst.Open strSql,Conn,1,1
		'Response.Write "���������ļ�¼��ĿΪ:" & rst.RecordCount
		If rst.Eof  Then
			'������������⴦��һ��
			If strMode = "3" Then
				strReturn = GetPhotoComment(rst,strUnit)
			Else
				If strMode = "2" Then
					'�ݹ�
					strReturn = ShowList("select photo_path,fileID,photo_Title,photo_name from oblog_album where TeamID=0 and (ishide=0 OR ishide IS NULL) and userid="&mUserId&" and userClassId=0  order by photoID desc","����Ƭ","4")
				Else
					strReturn=strReturn & "<ul>�޼�¼���������ݱ�����</ul>"
					rst.Close
					Set rst=Nothing
				End if
			End if
			ShowList = strReturn
			Exit Function
		End If
		G_P_AllRecords=rst.RecordCount
		'strReturn=strReturn & "������" & G_P_AllRecords & strUnit & "<br>"
		If G_P_This<1 Then
			G_P_This=1
		End If
		If (G_P_This-1)*G_P_PerMax>G_P_AllRecords Then
			If (G_P_AllRecords mod G_P_PerMax)=0 Then
				G_P_This= G_P_AllRecords \ G_P_PerMax
			Else
				G_P_This= G_P_AllRecords \ G_P_PerMax + 1
			End If
		End If
		If G_P_This=1 Then
			select Case strMode
					Case "0"
						strReturn = strReturn&ShowOnePage(rst)
						strReturn=strReturn & oblog.showpage(false,true,strUnit)
					Case "1"
						strReturn = ShowMessages(rst)
						strReturn="<h1 class=""message_title"">���԰���ҳ(<a href="""&blogdir&mUserPath&"/message."&f_ext&"#cmt"">ǩд����</a>)</h1>" & vbCrLf & strReturn & oblog.showpage(false,true,strUnit)
					Case "2","4"
						strReturn = strReturn&getPhotolist(rst)
						strReturn=strReturn & oblog.showpage(false,true,strUnit)
					Case "3"
						strReturn = strReturn&GetPhotoComment(rst,strUnit)
'						strReturn=strReturn & oblog.showpage(false,true,strUnit)
			End select
		Else
			If (G_P_This-1) * G_P_PerMax < G_P_AllRecords Then
				rst.move  (G_P_This-1) * G_P_PerMax
				'Dim bookmark
				'bookmark=rst.bookmark
				select Case strMode
					Case "0"
						strReturn = ShowOnePage(rst)
						strReturn=strReturn & oblog.showpage(false,true,strUnit)
					Case "1"
						strReturn = ShowMessages(rst)
						strReturn="<h1 class=""message_title"">���԰���ҳ(<a href="""&blogdir&mUserPath&"/message."&f_ext&"#cmt"">ǩд����</a>)</h1>" & vbCrLf & strReturn & oblog.showpage(false,true,strUnit)
					Case "2","4"
						strReturn = strReturn&getPhotolist(rst)
						strReturn=strReturn & oblog.showpage(false,true,strUnit)
					Case "3"
						strReturn = strReturn&GetPhotoComment(rst,strUnit)
'						strReturn=strReturn & oblog.showpage(false,true,strUnit)
				End select
			Else
				G_P_This=1
				select Case strMode
					Case "0"
						strReturn = ShowOnePage(rst)
						strReturn=strReturn & oblog.showpage(false,true,strUnit)
					Case "1"
						strReturn = ShowMessages(rst)
						strReturn="<h1 class=""message_title"">���԰���ҳ(<a href="""&blogdir&mUserPath&"/message."&f_ext&"#cmt"">ǩд����</a>)</h1>" & vbCrLf & strReturn & oblog.showpage(G_P_FileName,G_P_AllRecords,G_P_PerMax,false,true,strUnit)
					Case "2","4"
						strReturn = strReturn&getPhotolist(rst)
						strReturn=strReturn & oblog.showpage(false,true,strUnit)
					Case "3"
						strReturn = strReturn&GetPhotoComment(rst,strUnit)
'						strReturn=strReturn & oblog.showpage(false,true,strUnit)
				End select
			End If
		End If
		rst.Close
		Set rst=Nothing
		ShowList=strReturn
	End Function

	Private Function ShowOnePage(rst)
		Dim strBody,strContent,strTmp,rssubject,i,substr
		Dim strTopic,strLoginfo,strLogtext,strMore,strEmot,strAuthor,strAddtime,strTopictxt
		Set rssubject = oblog.execute("select subjectid,subjectname from oblog_subject where userid="&mUserid)
		While Not rssubject.EOF
			substr = substr & rssubject(0) & "!!??((" & rssubject(1) & "##))=="
			rssubject.movenext
		Wend
		substr = substr & "0!!??((ȫ����־##))=="
		i=0
		Do While Not rst.EOF
			if (mUsersublist=1 and id>0) or mUserIndexlist=1 then '�б���ʾ
				strBody="<li><a href="&mUserLogpath&rst("logfile")&" >"&oblog.filt_html(rst("topic"))&"</a>��"&oblog.filt_html(rst("author"))&" <span>"&rst("addtime")&"</span></li>"&vbcrlf
			else
				'If rst("face") = "0" Then
	'					strEmot = ""
	'				Else
	'					strEmot = "<img src="&blogurl&"images/face/" & rst("face") & ".gif />"
	'				End If
				If mUserNickName = "" Or IsNull(mUserNickName) Then
					strAuthor = mUserName
				Else
					strAuthor = mUserNickName
				End If

				If rst("authorid") <> mUserId Then strAuthor = rst("author")
				strAddtime = rst("addtime")
				strTopic = strEmot
				If rst("subjectid") > 0 Then
					strTopic = strTopic & "<a href=""" & mUserCmdpath & "cmd."&f_ext&"?uid="&mUserid&"&do=blogs&id=" & rst("subjectid") & """>[" & oblog.filt_html(getsubname(rst("subjectid"),substr)) & "]</a>"
				End If
				strTopictxt = "<a href=""" & mUserLogpath&rst("logfile") & """>" & oblog.filt_html(rst("topic")) & "</a>"
				If rst("isbest") = 1 Then strTopictxt = strTopictxt & "��<img src=" & blogurl & "images/jhinfo.gif >"
				Dim digg
				digg = ShowDigg
				digg = Replace(digg,"$diggnum$",OB_IIF(rst("DIGGNum"),0))
				digg = Replace(digg,"$logid$",rst("logid"))
				digg = Replace(digg,"$showmsg$","�Ƽ�")

				strTopic = strTopic & strTopictxt
				If rst("istop") = 1 Then strTopictxt = "[�ö�]" & strTopictxt
				strLoginfo = strAuthor & " ������ " & strAddtime
				strMore = "<a href=""" & mUserLogpath&rst("logfile") & """>�Ķ�ȫ��("&rst("iis")&")</a>"
				strMore = strMore & " | <a href=""" & mUserLogpath & rst("logfile") & "#cmt"">�ظ�("&rst("commentnum")&")</a> | <a href=""javascript:void(null);"" onclick=""ListMenu('"&rst("logid")&"')""><span id=""a_"&rst("logid")&""">��ӳ����</span></a><span id=""menu_"&rst("logid")&"""></span>"
				strMore = strMore & " | <a href=""" & blogurl & "showtb.asp?id=" & rst("logid") & """ target=""_blank"">����ͨ��("&rst("trackbacknum")&")</a>"
				'������ǰ����
				If rst("ishide") = 1 Then strTmp = "����־Ϊ������־�������ѿɼ���<a href='" & blogurl & "more.asp?id=" & rst("logid") & "'>���������֤ҳ��</a>��"
				If rst("ispassword") <> "" Then strTmp = "<form method='post' action='" & blogurl & "more.asp?id=" & rst("logid") & "' target='_blank'>��������־�������룺<input type=""password"" size=""15"" name=""password"" />  <input type=""submit"" value=""�ύ""></form>"
				If rst("passcheck") = 0 Then strTmp = "����־��Ҫ����Ա��˺�ſɼ���"
				If strTmp <> "" Then
					strLogtext = strTmp
					strTmp = ""
				Else
					If rst("Abstract")="" Or IsNull(rst("Abstract"))  then
						strLogtext = rst("logtext")
						strLogtext = trimlog(strLogtext, rst("showword"))
						'If Left(strLogtext, 7) = "#isubb#" Then
							'strLogtext = UBBCode(strLogtext, 1)
							'strLogtext = Replace(strLogtext, Chr(10), "<br /> ")

						'End If
						strLogtext = Replace(strLogtext, "#isubb#", "")
						strLogtext = filtimg(strLogtext)
						If oblog.cacheConfig(45)=1 Then strLogtext = profilthtm(strLogtext)
						strLogtext = "<span id=""ob_logd"& rst("logid") &""">"&digg&"</span>"&strLogtext
					Else
						strLogtext = "<span id=""ob_logd"& rst("logid") &""">"&digg&"</span>"&rst("Abstract")
					End if
				End If
				 strLogtext=oblog.filt_badword(UBBCode(strLogtext,1))
				 '��ʹ�����·��ʱ���滻Ϊ����·��
				 'if oblog.CacheConfig(55) = "1" then
					'	strLogtext=filtskinpath(strLogtext)
				 'end if
				 strlogn = strlogn & "$" & rst("logid")
				 strBody = Replace(mUserSkinLog, "$show_topic$", strTopic)
				 strBody = Replace(strBody, "$show_loginfo$", strLoginfo)
				 strBody = Replace(strBody, "$show_logtext$", strLogtext)
				 strBody = Replace(strBody, "$show_more$", strMore)
				 strBody = Replace(strBody, "$show_emot$", strEmot)
				 strBody = Replace(strBody, "$show_author$", strAuthor)
				 strBody = Replace(strBody, "$show_addtime$", strAddtime)
				 strBody = Replace(strBody, "$show_topictxt$", strTopictxt)
				 strBody = Replace(strBody, "$show_blogzhai$", "")
				 strBody = Replace(strBody, "$show_blogtag$", "")
				 'show_logmore = show_logmore & strBody
			 end if
			 strContent = strContent & VBCRLF & strBody
			 rst.movenext
			 i=i+1
			 if i>=G_P_PerMax then exit do
		  Loop
		  set rssubject=nothing
		  ShowOnePage=strContent
		  if (mUsersublist=1 and id>0) or mUserIndexlist=1then
			ShowOnePage="<div id=""subject_index""><ul>"&oblog.filt_html(getsubname(id,substr))&ShowOnePage&"</ul></div>"
		  end if
	End Function

	Public Function ShowMessages(rst)
		Dim strtopic, stremot, straddtime, strlogtext, strauthor, strloginfo, strmore, strMessage, strtopictxt, strContent
		Dim homepage_str, user_filepath,i
		If Not rst.EOF Then
			Do While Not rst.EOF
				If IsNull(rst("homepage")) Then
					homepage_str = "������ҳ"
				Else
					If Trim(Replace(rst("homepage"), "http://", "")) = "" Then
						homepage_str = "������ҳ"
					Else
						homepage_str = "<a href=""" &blogurl&"go.asp?url=" & oblog.filt_html(rst("homepage")) & """ target=""_blank"">������ҳ</a>"
					End If
				End If
				strtopic = oblog.filt_html(rst("messagetopic")) & "<a name='" & rst("messageid") & "'></a>"
				If rst("isguest") = 1 Then
					strauthor = oblog.filt_html(rst("message_user")) & "(�ο�)"
				Else
					strauthor = oblog.filt_html(rst("message_user"))
				End If
				straddtime = rst("addtime")
				strtopictxt = strtopic
				strloginfo = strauthor & "����������" & straddtime
				'strlogtext = oblog.Ubb_Comment(rst("message"))
				If rst("ubbedit")= 2 Then
					strlogtext = oblog.FilterUbbFlash(filtscript(rst("message")))
				Else
					strlogtext = oblog.Ubb_Comment(rst("message"))
				End if
				strmore = homepage_str & " | <a href='"&blogurl&"user_messages.asp?action=modify&re=true&id=" & rst("messageid") & "'>�ظ�</a>"
				strmore = strmore & " | <a href=""" & blogurl & "user_messages.asp?action=del&id=" & rst("messageid") & """  target=""_blank"">ɾ��</a>"
				if rst("ishide")=1 then
					strtopictxt="���Ļ�"
					strtopic="���Ļ�"
					strlogtext="������Ϊ���Ļ���"
					strmore=Replace(strmore,"�ظ�","�鿴")
				end if
				strMessage = Replace(mUserSkinLog, "$show_topic$", strtopic)
				strMessage = Replace(strMessage, "$show_loginfo$", strloginfo)
				strMessage = Replace(strMessage, "$show_logtext$", strlogtext)
				strMessage = Replace(strMessage, "$show_more$", strmore)
				strMessage = Replace(strMessage, "$show_emot$", "")
				strMessage = Replace(strMessage, "$show_author$", strauthor)
				strMessage = Replace(strMessage, "$show_addtime$", straddtime)
				strMessage = Replace(strMessage, "$show_topictxt$", strtopictxt)
				strMessage = Replace(strMessage, "$show_blogtag$", "")
				strMessage = Replace(strMessage, "$show_blogzhai$", "")
				strContent = strContent & strMessage
				rst.movenext
				i=i+1
				If i>=G_P_PerMax Then Exit Do
			Loop
		Else
			strContent = "��������"
		End If
		ShowMessages=strContent
	End Function

	'��ȡ�û���Ϣ
	Private Function GetUserInfo()
		Dim rst,rst1,ustr
		Set rst=oBlog.Execute("select * From oBlog_User Where UserId=" & mUserId)
		If rst.Eof Then
			Set rst = Nothing
'			GetUserInfo= "������û����"
			Exit Function
		Else
			'�ж��Ƿ���վ����
			if rst("blog_password")<>""  and Request.Cookies(cookies_name)("blog_pwd_"&mUserId)<>rst("blog_password") then
				set rst=nothing
				Response.Write "window.location='"&blogurl&"chkblogpassword.asp?userid="&mUserId&"';"
				Response.End()
			end if
			mUserFolder=rst("user_folder")
			mUserPath=rst("user_dir")&"/"&rst("user_folder")
			mBlogName=rst("blogname")
			mUserName=rst("username")
			mUserNickName=rst("nickname")
			G_P_PerMax=rst("user_showlog_num")
			mUserPhotoRow=rst("user_photorow_num")
			ustr=rst("user_info")
			mUserIndexlist=rst("indexlist")
			mUserIcon1 = rst ("user_icon1")
			if ustr="" or isnull(ustr) then
				mUsersublist=0
			else
				ustr=split(ustr,"$")
				if ustr(0)<>"" then mUsersublist=cint(ustr(0)) else mUsersublist=0
			end if
			if mUsersublist=1 and id>0 then G_P_PerMax=40 '�б�ģʽ����50��
			if mUserPhotoRow<=0 or isnull(mUserPhotoRow) then mUserPhotoRow=4
			If IsNull(rst("user_skin_showlog")) OR rst("user_skin_showlog")="" Then
				Set rst1 = oBlog.Execute("select skinshowlog from oBlog_userskin where isdefault=1")
				If Not rst1.EOF Then
					mUserSkinLog = rst1("skinshowlog")
					Set rst1 = Nothing
				Else
					Set rst1 = Nothing
					Set rs = Nothing
					Response.Write ("ģ�����")
					Response.End
				End If
			Else
				mUserSkinLog=rst("user_skin_showlog")
			End If
			if true_domain=1 then
				mUserCmdpath="/"
				mUserLogpath=""
			else
				mUserCmdpath=blogdir&mUserPath&"/"
				mUserLogpath=blogdir
			end if
			'mUserSkinLog=filtskinpath(mUserSkinLog)
		End If
		Set rst=Nothing
	End Function

	Function getPhotolist(rsPhoto)
		Dim i,bstr,n,fso,sReturn
		Dim title,imgsrc
		Dim goUrl,rsSubject,substr,subjectname
		Set rsSubject = oblog.execute("select subjectid,subjectname from oblog_subject where subjecttype = 1 AND userid="&mUserid)
		While Not rsSubject.EOF
			substr = substr & rsSubject(0) & "!!??((" & rsSubject(1) & "##))=="
			rsSubject.movenext
		Wend
'		OB_DEBUG substr,1
		Set rsSubject = Nothing
'		Set fso = Server.CreateObject(oblog.CacheCompont(1))
		'�����������
		If mUserPhotoRow > rsPhoto.RecordCount Then mUserPhotoRow = rsPhoto.RecordCount
		If ID = 0 Then
			sReturn=vbcrlf & "<table width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""1""><tbody>"& vbcrlf
			Do While not rsPhoto.eof
				sReturn=sReturn&"<tr>"& vbcrlf
				For n=1 to mUserPhotoRow
					if rsPhoto.eof then
'						sReturn=sReturn&"<td width=""25%""></td>"& vbcrlf
					Else
						subjectname = oblog.filt_html(getsubname(rsPhoto(1),substr))
						goUrl = mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=album&id="&rsPhoto(1)
						title = "<div class=""photo_album_name""><a href="""&goUrl&""" title=""��᣺"&subjectname&""">"&subjectname&"</a></div><div class=""photo_album_num"">��Ƭ����"&rsPhoto("subjectlognum")&"</div>"
						imgsrc = ProIco(rsPhoto(0),4)
						'imgsrc=Replace(imgsrc,right(imgsrc,3),"jpg")
						'imgsrc=Replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
						'if  not fso.FileExists(Server.MapPath(imgsrc)) then
							'imgsrc=blogurl&rsPhoto(0)
						'End if
						If oblog.CacheConfig(67) = "1" Then
							If Left(LCase(imgsrc),7)="http://" Then
							imgsrc=imgsrc
							Else
							imgsrc = "attachment.asp?path="&imgsrc
							End If
						End If

						sReturn=sReturn&"<td align='center'><div class=""photo_album"" style=""width:130px;height:160px;overflow:hidden;margin: 8px 0;padding:10px 0 0 0;background:url("&blogurl&"Images/photo_album.gif) no-repeat left top;""><div class=""photo_ico""><table><tr><td><a href="""&goUrl&"""><img src='"&imgsrc&"' style=""vertical-align:middle;max-width: 100px; max-height: 100px; width: expression(this.width >100 && this.height < this.width ? 100: true); height: expression(this.height > 100 ? 100: true);"" align=""absmiddle"" /></a></td></tr></table></div>"&title&"</div></td>"& vbcrlf
						i=i+1
						rsPhoto.movenext
					End if
				Next
				sReturn=sReturn&"</tr>"& vbcrlf
				if i>=G_P_PerMax then exit do
			Loop
			If id = 0 And 1=2 Then
				Dim trs,rsPic,DefaultPic
				Set trs = Oblog.Execute ("SELECT COUNT(photoID) FROM oblog_album WHERE TeamID=0 AND (ishide=0 OR ishide IS NULL) AND userid="&mUserId&" AND userClassId=0 ")
				Set rsPic = Oblog.Execute ("SELECT TOP 1 photo_path FROM oblog_album WHERE TeamID=0 AND (ishide=0 OR ishide IS NULL) AND userid="&mUserId&" AND userClassId=0 ")
				If Not rsPic.Eof Then
					DefaultPic = rsPic(0)
					rsPic.Close
					Set rsPic = Nothing
				End if
				goUrl = mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=album&id=-1"
				title = "<div class=""photo_album_name""><a href="""&goUrl&""" title=""��᣺δ����"">δ����</a></div><div class=""photo_album_num"">��Ƭ����"&TRS(0)&"</div>"
				sReturn=sReturn&"<tr><td align='center'><div class="""&album_ClassName&"""><div class=""photo_ico""><table><tr><td><a href="""&goUrl&"""><img src='"&ProIco(DefaultPic,4)&"' align=""absmiddle"" /></a></td></tr></table></div>"&title&"</div></td></tr>"& vbcrlf
				Set TRS = Nothing
			End if
			sReturn=sReturn&"</tbody></table>"	& VBCRLF
		Else
		'����������Ƭ
			sReturn=vbcrlf & "<table width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""1""><tbody>"& vbcrlf
			Do While not rsPhoto.eof
				sReturn=sReturn&"<tr>"& vbcrlf
				For n=1 to mUserPhotoRow
					if rsPhoto.eof then
'						sReturn=sReturn&"<td width=""25%""></td>"& vbcrlf
					Else
						Oblog.Execute ("UPDATE oblog_subject SET views = views + 1 WHERE subjectid="&id)
						goUrl = mUserCmdpath&"cmd."&f_ext&"?do=photocomment&fileid="&rsPhoto(1)&"&uid="&mUserid
						title="<div class=""photo_name""><a href="""&goUrl&""" title="""&ob_IIF(rsPhoto(2),"�ޱ���")&""">"&ob_IIF(rsPhoto(2),"�ޱ���")&"</a></div>"
						imgsrc = ProIco(rsPhoto(0),4)
						'imgsrc=Replace(imgsrc,right(imgsrc,3),"jpg")
						'imgsrc=Replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
						'if  not fso.FileExists(Server.MapPath(imgsrc)) then
							'imgsrc=blogurl&rsPhoto(0)
						'End if
						If oblog.CacheConfig(67) = "1" Then

							If Left(LCase(imgsrc),7)="http://" Then
							imgsrc=imgsrc
							Else
							imgsrc = "attachment.asp?path="&imgsrc
							End If
						End If
						Dim lightboxstr
						If CBool(Islightbox) Then
						lightboxstr=" href="""&imgsrc&""" rel=""lightbox[roadtrip]"" "
						Else
						lightboxstr=" href="""&goUrl&""" "
						End If
						sReturn=sReturn&"<td align='center'><div class=""photo_album_list""><div class=""photo_ico""><table><tr><td><a "&lightboxstr&"><img src='"&imgsrc&"' style=""vertical-align:middle;max-width: 100px; max-height: 100px; width: expression(this.width >100 && this.height < this.width ? 100: true); height: expression(this.height > 100 ? 100: true);"" align=""absmiddle"" /></a></td></tr></table></div>"&title&"</div></td>"& vbcrlf
						i=i+1
						rsPhoto.movenext
					End if
				Next
				sReturn=sReturn&"</tr>"& vbcrlf
				if i>=G_P_PerMax then exit do
			Loop
			sReturn=sReturn&"</tbody></table>"	& VBCRLF
		End if
'		Set fso=nothing
		getPhotolist=sReturn
	End Function

	'��ȡ�û�����
	Function GetUserClasses(typestr)
		Dim rst,sReturn
		Set rst=conn.Execute("select * From oblog_subject Where subjecttype=1 AND (ishide = 0 OR ishide IS NULL) and userid="&mUserid&" order by ordernum")
		If rst.Eof Then
			sReturn=""
		Else
			Do While Not rst.Eof
				sReturn=sReturn&"<option value="&rst("subjectid")&">" & rst("subjectname") & "</option>" & VBCRLF
				rst.Movenext
			Loop
			sReturn = "<option value="""">��ѡ����Ƭ����</option><option value='0'>���з���</option>" & VBCRLF & sReturn
			sReturn = sReturn &"<option value='-1'>δ����</option>" & VBCRLF
			sReturn="<select name=classid onchange=""javascript:window.location='"&mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do="&typestr&"&id='+this.options[this.selectedIndex].value;"">" & VBCRLF & sReturn & "</select>"
		End If
		rst.Close
		Set rst=Nothing
'		sReturn=sReturn&"��<a href=""#"" onclick=""window.open('"&strPlayerUrl&"','_photo','height=500, width=480, top=100, left=400, toolbar=no, menubar=no, scrollbars=no, resizable=yes,status=no')"">�����Զ�����</a>" & VBCRLF
		If CBool(Is_Sqldata) Then
			sReturn=sReturn&"  <a href='"&blogurl&"PhotoViewer.asp?uid="&mUserid&"&do=flash' target=""_blank"">Flashȫ����ʽ���</a>"
		Else
			sReturn=sReturn&"  <a href='"&mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=flash'>Flash��ʽ���</a>"
		End If
		sReturn=sReturn&"  <a href='"&mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=album&id=-2'>��ʾ������Ƭ</a>"
		GetUserClasses = sReturn
	End Function


	function showinfo()
		dim rs,str,c0,c1,c2
		select case Trim(Request("infotype"))
		case "1"
			str=str&"<ul class=""ob_user_friend"">"
			set rs=oblog.execute("select a.username,a.nickname,a.user_icon1,a.userid from oblog_user a,oblog_friend  b where a.userid=b.friendid and b.userid ="&mUserid&" and b.isblack=0 order by b.id desc")
			while not rs.eof
				str=str&"<li><a href=" &blogurl & "go.asp?userid="&rs(3)&" target=_blank><img src=""" & ProIco(rs(2),1) & """ class=""ob_face_info"" /><br />"&OB_IIF(rs(1),rs(0))&"</a></li>"&vbcrlf
				rs.movenext
			wend
			str=str&"</ul>"& vbcrlf
			c1=" class='nowselect' "
		case "2"
			str=str&"<ul class=""ob_user_group"">"
			set rs=oblog.execute("select a.t_name,a.teamid,a.t_ico from oblog_team a,oblog_teamusers  b where a.teamid=b.teamid and a.istate=3 and (b.state=3 or b.state=5 ) and userid ="&mUserid)
			while not rs.eof
				str=str&"<li><a href=" &blogurl & "group.asp?gid="&rs(1)&" target=""_blank""><img src=""" & ProIco(rs(2),2) & """ class=""group_logo_info"" /><br />"&oblog.filt_html(left(rs(0),18))&"</a></li>"&vbcrlf
				rs.movenext
			wend
			str=str&"</ul>"& vbcrlf
			c2=" class='nowselect' "
		case else
			set rs=oblog.execute("select * from oblog_user where userid="&mUserid)
			if not rs.eof then
				str=str&"<ul class=""ob_user_info""><img src=""" & ProIco(rs("user_icon1"),1) & """ class=""ob_face_info"" /><li>�û�����"&rs("username")&"</li><li>�ǡ��ƣ�"&rs("nickname")&"</li><li>�ԡ���"&ob_IIF2(rs("sex")=1,"��","Ů")&"</li><li>�桡����"&rs("truename")&"</li><li>���ڵأ�"&rs("province")&rs("city")&"</li><li>�����գ�"&rs("birthday")&"</li><li>ְ��ҵ��"&rs("job")&"</li><li>MSN��"&rs("msn")&"</li><li>Q Q��"&rs("qq")&"</li><li>�ء�ַ��"&rs("address")&"</li><li>�򡡽飺"&oblog.filt_html(rs("siteinfo"))&"</li></ul>"& vbcrlf
			end if
			c0=" class='nowselect' "
		end select
		showinfo="<div id=""ob_userinfo""><ul class=""top""><li"&c0&"><a href='"&mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=info'>��ϸ����</a></li><li"&c1&"><a href='"&mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=info&infotype=1'>����</a></li><li"&c2&"><a href='"&mUserCmdpath&"cmd."&f_ext&"?uid="&mUserid&"&do=info&infotype=2'>" &oblog.CacheConfig(69)& "</a></li></ul>"&str&"</div>"
		set rs=nothing
	end function

	Function GetPhotoComment(rsPComment,strUnit)
		Dim trs,i
		Dim sPInfo,sPComment,imgsrc
		Dim show_topic,show_emot,show_addtime,show_logtext,show_author,show_loginfo,show_more,show_logcyc,show_topictxt,show_logmore,show,commentasc,faceurl
		dim homepage_str,commentid,strtmp
		Dim commenttopic
		Dim allsub,substr,rsSubject
		Set rsSubject=Server.CreateObject("Adodb.Recordset")
		rsSubject.Open "select subjectid,subjectname From oblog_subject Where userid=" & mUserid & " And subjecttype=1",conn,1,3
		If Not rsSubject.Eof Then
			Do While Not rsSubject.Eof
				allsub=allsub&rsSubject(0)&"!!??(("&rsSubject(1)&"##))=="
				rsSubject.movenext
			Loop
		End If
		Set rsSubject = Nothing
		Set trs = oblog.Execute ("select PHOTO_title,PHOTO_readme,PHOTO_path,a.addtime,b.nickname,b.username,a.views,a.isencomment,a.userClassId,a.commentnum FROM oblog_album a INNER JOIN oblog_user b ON a.userid = b.userid WHERE (a.ishide = 0 OR A.ishide IS NULL) AND a.TeamID = 0 AND a.fileid="&fileid)
		If TRS.EOF Then
			GetPhotoComment = "����Ƭ�����ڣ����߱����ء�"
			trs.Close
			Set trs = Nothing
			Exit Function
		Else
			Dim thisSubName
			thiSsubName=getsubname(trs("userClassId"),allsub)
			If thiSsubName="����" Then thiSsubName =  ""
			thiSsubName = "["&thiSsubName&"]"
			Oblog.Execute ("UPDATE oblog_album SET views = views + 1 WHERE fileid="&fileid)
			If oblog.CacheConfig(67) = "1" Then
				imgsrc = "attachment.asp?path="&trs("PHOTO_path")
			Else
				imgsrc = ProIco(trs(2),4)
			End If
			sPInfo = Replace(mUserSkinLog, "$show_topic$", ob_IIF(thiSsubName&trs(0),"�ޱ���"))
			sPInfo = Replace(sPInfo, "$show_loginfo$", "")
			sPInfo = Replace(sPInfo, "$show_logtext$", "<br/><img src="&imgsrc&" onclick=""javascript:window.open(this.src);"" style=""CURSOR: pointer"" onload=""rsimg(this,400);"" /><br />"&ob_IIF(trs(1),"�޼��"))
			sPInfo = Replace(sPInfo, "$show_more$", "<a href=""#"">�鿴("&OB_IIF(trs(6),0)&")</a>&nbsp;|&nbsp;<a href=""#cmt"">����("&OB_IIF(trs(9),0)&")</a>")
			sPInfo = Replace(sPInfo, "$show_emot$", "")
			sPInfo = Replace(sPInfo, "$show_author$", OB_IIF(trs(4),trs(5)))
			sPInfo = Replace(sPInfo, "$show_addtime$",trs(3))
			sPInfo = Replace(sPInfo, "$show_topictxt$", "")
			sPInfo = Replace(sPInfo, "$show_blogtag$", "")
			sPInfo = Replace(sPInfo, "$show_blogzhai$", "")
		End If
		commenttopic = "Re:" & ob_IIF(trs(0),"�ޱ���")
'		GetPhotoComment = sPInfo & "<div style=font-size:14px;font-weight:600>�����б�</div>"
		GetPhotoComment = sPInfo
		If rsPComment.EOF Then
'			GetPhotoComment = GetPhotoComment & "�������������"
			rsPComment.Close
			Set rsPComment = Nothing
		Else
			i = 0
			Do While Not rsPComment.EOF
				if isnull(rsPComment(1)) then
					homepage_str="������ҳ"
				else
					if Trim(Replace(rsPComment(1),"http://",""))="" then
						homepage_str="������ҳ"
					else
						homepage_str="<a href="""&oblog.filt_html(rsPComment(1))&""" target=""_blank"">������ҳ</a>"
					end if
				end If
				commentid=rsPComment(4)
				show_topic=oblog.filt_html(rsPComment(2))&"<a name='"&rsPComment(4)&"'></a>"
				if rsPComment(6)=1 then
					show_author="<span id=""n_"&commentid&""">"&oblog.filt_html(rsPComment(0))&"(�ο�)</span>"
					faceurl=blogurl&"images/ico_default.gif"
				else
					show_author="<span id=""n_"&commentid&""">"&oblog.filt_html(rsPComment(0))&"</span>"
					Dim rsUser
					Set rsUser = oblog.Execute ("SELECT user_icon1 FROM oblog_user WHERE username = '"&rsPComment("comment_user")&"'")
					If Not rsUser.Eof Then
						faceurl = ProIco (rsUser(0),1)
					Else
						faceurl=blogurl&"images/ico_default.gif"
					End if
				end If
				faceurl="<img class=""ob_face"" src="""&faceurl&""" width=""48"" height=""48"" align=""absmiddle"" />"
				faceurl=Replace(homepage_str,"������ҳ",faceurl)
				show_addtime="<span id=""t_"&commentid&""">"&rsPComment(5)&"</span>"
				show_topictxt=show_topic
				show_loginfo=show_author&"����������"&show_addtime
				sPComment = faceurl &"<span id=""c_" & commentid & """>"
				sPComment = sPComment & oblog.Ubb_Comment(rsPComment(3))
				sPComment = sPComment &"</span>"
				show_more=homepage_str&" | <a href=""javascript:reply_quote('"&commentid&"')"" >����</a> | <a href=""#top"">����</a>"
				show_more=show_more&" | <a href=""user_comments.asp?action=del&id="&commentid&"""  target=""_blank"">ɾ��</a>"
				show_logcyc=Replace(mUserSkinLog,"$show_topic$",show_topic)
				show_logcyc=Replace(show_logcyc,"$show_loginfo$",show_loginfo)
				show_logcyc=Replace(show_logcyc,"$show_logtext$",sPComment)
				show_logcyc=Replace(show_logcyc,"$show_more$",show_more)
				show_logcyc=Replace(show_logcyc,"$show_emot$","")
				show_logcyc=Replace(show_logcyc,"$show_author$",show_author)
				show_logcyc=Replace(show_logcyc,"$show_addtime$",show_addtime)
				show_logcyc=Replace(show_logcyc,"$show_topictxt$",show_topictxt)
				show_logmore=show_logmore&show_logcyc
				show_logmore = Replace(show_logmore, "$show_blogtag$", "")
				show_logmore = Replace(show_logmore, "$show_blogzhai$", "")
				rsPComment.MoveNext
				i = i + 1
				If i>=G_P_PerMax Then Exit Do
			Loop
			show_logmore = show_logmore &oblog.showpage(false,true,strUnit)
		End If
		If trs("isencomment") = "1" Then
			Dim strguest
			If oblog.cacheConfig(27) = 1 Then strguest = "(�ο�������������)" Else strguest = ""
			show_logmore = filt_inc(show_logmore)
			show_logmore = show_logmore & vbCrLf & "<div id=""form_comment"">" & vbCrLf
			show_logmore = show_logmore & "	#ad_usercomment#<a name=""cmt""></a><div class=""title"">�������ۣ�</div>" & vbCrLf
			show_logmore = show_logmore & "	<form action=""" & blogurl & "SaveAlbumComment.asp?fileid=" & Fileid & """ method=""post"" name=""commentform"" id=""commentform"" onSubmit=""return Verifycomment()"">" & vbCrLf
			show_logmore = show_logmore & "		<div class=""d1""><label>�ǳƣ�<input name=""UserName"" type=""text"" id=""UserName"" size=""20"" maxlength=""20"" value="""" /></label></div>" & vbCrLf
			show_logmore = show_logmore & "		<div class=""d2""><label>���룺<input name=""Password"" type=""password"" id=""Password"" size=""20"" maxlength=""20"" value="""" /> " & strguest & "</label></div>" & vbCrLf
			show_logmore = show_logmore & "		<div class=""d3""><label>��ҳ��<input name=""homepage"" type=""text"" id=""homepage"" size=""42"" maxlength=""50"" value=""http://"" /></label></div>" & vbCrLf
			show_logmore = show_logmore & "		<div class=""d4""><label>���⣺<input name=""commenttopic"" type=""text"" id=""commenttopic"" size=""42"" maxlength=""50"" value=""" & commenttopic & """ /></label></div>" & vbCrLf
			show_logmore = show_logmore & "		<div class=""d5"">" & vbCrLf
			show_logmore = show_logmore & "			<input type=""hidden"" name=""edit"" id=""edit"" value="""" />" & vbCrLf
			show_logmore = show_logmore & "			<div id=""oblog_edit"">"& oblog.CacheConfig(41)&"</div>" & vbCrLf
			show_logmore = show_logmore & "		</div>" & vbCrLf
			show_logmore = show_logmore & "		<div class=""d6""><span id=""ob_code""></span><input type=""submit"" value=""&nbsp;��&nbsp;��&nbsp;"" onclick='oblog_edittext.createTextRange().execCommand(""Copy"");'></div>" & vbCrLf
			show_logmore = show_logmore & "	</form>" & vbCrLf
			show_logmore = show_logmore & "</div>" & vbCrLf
			show_logmore = Replace(show_logmore, "#ad_usercomment#", "<div id=""ad_usercomment""></div>")
		End if
		GetPhotoComment = GetPhotoComment & show_logmore
	End Function
End Class
%>
