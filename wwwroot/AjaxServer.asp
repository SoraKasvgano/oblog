<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/class_trackback.asp"-->
<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.cachecontrol = "no-cache"
'Oblog4.0 AJAX Server
'------------------------------------------------
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
Dim Action,tName
action=LCase(Request("action"))
tName="��־"
select Case action
	Case "get_draft"
		Call get_draft()
	Case "savelog"
		Call savelog()
	Case "getfeedlist"
		Call getfeedlist()
	Case "getpm"
		Call getpm()
	Case "vote"
		Call SaveVote
	Case "digglog"
		Call digglog()
	Case "savereport"
		Call SaveReport()
End select

Sub SaveVote()
	Dim sValue,logid,rs,Scores,targetUserid
	sValue=Request("v")
	logid=FilterIds(Int(Request("logid")))
	If sValue<>"1" Then sValue=0
	'-------------------------------
	'1.���е�¼���
	'-------------------------------
	If Not oblog.checkuserlogined() Then
		Response.Write "�����¼����ܽ��д˲���"
		Response.End
	Else
		If oblog.l_ulevel=6 then
			Response.Write "�����ʺŻ�û��ͨ����ˣ����ܽ��д˲���"
			Response.End
		End If
	End If
	'-------------------------------
	'2.����û������Ƿ��㹻
	'-------------------------------
	'If oblog.CheckScore(oblog.CacheScores(20)) Then
	'	Response.Write "�ò�����Ҫ " & oblog.CacheScores(20) & " ���֣����Ļ��ֲ���"
	'	Response.End
	'End If
	'-------------------------------
	'3.���Ŀ������
	'-------------------------------
	'����ƽ��ά��
	'�ڷ���״̬�£�����һζ�ļ��٣�����ֻ���ٵ�������־���л��ֿ۳���Ϊֹ
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "select scores,userid From oblog_log Where logid=" & logid ,conn,1,3
	If rs.Eof Then
		rs.Close
		Set rs=Nothing
		Response.Write "Ŀ�����²�����"
		Response.End
	End If
	targetUserid=rs(1)
	Scores=OB_IIF(rs(0),0)
	If targetUserid=oblog.l_uid Then
		rs.Close
		Set rs=Nothing
		Response.Write "���Լ����ܸ����Լ������½��д˲���"
		Response.End
	End If
	rs.Close
	'-------------------------------
	'4.����Ƿ��ѱ�̬��,���û������м�¼
	'-------------------------------
	rs.Open "select * From oblog_logvotes Where logid=" & logid & " And userid=" & oblog.l_uid,conn,1,3
	If Not rs.Eof Then
		Response.Write "��֮ǰ�Ѿ���̬Ϊ "
		If rs("vote")=1 Then
			Response.Write C_Vote_Action1
		Else
			Response.Write C_Vote_Action2
		End If
		rs.Close
		Set rs=Nothing
		Response.End
	End If
	rs.AddNew
	rs("logid")=logid
	rs("userid")=oblog.l_uid
	rs("vote")=sValue
	rs("addtime")=oblog.ServerDate(Now)
	rs("addip")=oblog.userIp
	rs.Update
	rs.Close
	'-------------------------------
	'5.���л��ֲ��� OS*#046
	'Ŀ���û�����+/-
	'-------------------------------
	'�����ǰ����>���۳����֣������ֵ/�����ڼ���ʱ
	If oblog.CacheScores(20)<>"" Then
		If Scores> Int (oblog.CacheScores(20)) Then
			Scores=oblog.CacheScores(20)
		End If
		If sValue="1" Then
			oblog.execute("Update oblog_log Set vote1=vote1+1,scores=scores+" & oblog.CacheScores(20)&" Where logid=" & logid)
			oblog.execute("Update oblog_user Set scores=scores+" & oblog.CacheScores(20)&" Where userid=" & targetUserid)
		Else
			oblog.execute("Update oblog_log Set vote0=vote0+1,scores=scores-" & Scores &"  Where logid=" & logid)
			oblog.execute("Update oblog_user Set scores=scores-" & oblog.CacheScores(20)&" Where userid=" & targetUserid)
		End If
		'-------------------------------
		'6.�۳���Դ�û����� OS*#046
		'-------------------------------
		oblog.execute("Update oblog_user Set scores=scores-" & oblog.CacheScores(20)&" Where userid=" & oblog.l_uid)
	End If
	Response.Write "�������"
	Dim blog
	set blog=new class_blog
	blog.userid=targetUserid
	Server.ScriptTimeOut=99999
	blog.update_log logid,3
	'Call blog.CreateFunctionPage
	set blog=nothing
End Sub


'��ȡfeed�б�
sub getfeedlist()
	if not oblog.checkuserlogined() then
		exit sub
	end if
	dim rsSubject,rs,str,t,m,ostr,isupdate,ajax,n
	Set rsSubject = oblog.Execute("select subjectid,subjectname from oblog_subject where userid=" & oblog.l_uId & " And subjecttype=3 order by ordernum")
	set rs=oblog.execute("select * from oblog_myurl where subjectid>0 and userid="&oblog.l_uid&" order by subjectid desc")
	n=0
	while not rsSubject.eof
		'if m=1 then ostr="</ol>" else ostr=""
		str=str&ostr&"<li id='su_"&n&"' class='open' onClick=""if(ol_"&n&".style.display == 'none'){ol_"&n&".style.display = '';}else{ol_"&n&".style.display = 'none';};su_click(document.getElementById('su_"&n&"'));""><a href='#' title="""&rsSubject("subjectname")&""" >"&rsSubject("subjectname")&"</a></li><ol id='ol_"&n&"'>"
		rs.Filter = "subjectid = " &rsSubject("subjectid")
		while not rs.eof
			if rs("isupdate")=1 then isupdate="class='isupdate'" else isupdate="class=noupdate"
			str=str&"<li id='now701' "&isupdate&" onclick=""this.className='noupdate'""><a  href='user_url.asp?action=read&feedurl="&rs("url")&"&encodeing="&rs("encodeing")&"&title="&Server.UrlEncode(rs("title"))&"&mainuserid="&rs("mainuserid")&"' onclick=""go_cmdurl('�ҵĶ���',this)"" target='content3' title="""&rs("title")&""">"&rs("title")&"</a></li>"
			rs.movenext
		Wend
		str=str&"</ol>"
		rsSubject.movenext
		n=n+1
	wend
	set rs=oblog.execute("select * from oblog_myurl where subjectid=0 and userid="&oblog.l_uid)
	if  not rs.eof then
		str=str&ostr&"<li id='su_999' class='open' onClick=""if(ol_no.style.display == 'none'){ol_no.style.display = '';}else{ol_no.style.display = 'none';}su_click(document.getElementById('su_999'));""><a href='#' title=""δ����"">δ����</a></li><ol id='ol_no'>"
		while not rs.eof
			if rs("isupdate")=1 then isupdate="class='isupdate'" else isupdate="class=noupdate"
			str=str&"<li id='now701' "&isupdate&"><a href='user_url.asp?action=read&feedurl="&rs("url")&"&encodeing="&rs("encodeing")&"&title="&Server.UrlEncode(Ob_iif(rs("title"),"δ����"))&"&mainuserid="&rs("mainuserid")&"' onclick=""go_cmdurl('�ҵĶ���',this)"" target='content3' title="""&rs("title")&""">"&rs("title")&"</a></li>"
			rs.movenext
		wend
	end if
	set rs=nothing
	set rsSubject=nothing
	'Response.Write(str)
	'if str="" then str="<span style='margin:20px;'>���޶���(<a href='user_url.asp' target='content3'>����</a>)</span>"
	str="<ol class=""option""><li class=""t1""><a id='active701' href='user_url.asp?action=add' onclick=""go_cmdurl('���Ӷ���',this)"" target='content3' title='���Ӷ���'>����</a></li><li class=""t1""><a id='active702' href='user_url.asp' onclick=""go_cmdurl('���Ĺ���',this);"" target='content3' title='���Ĺ���'>����</a></li><li class=""t1""><a id='active703' href='user_subject.asp?t=3' onclick=""go_cmdurl('���ķ���',this)"" target='content3' title='���ķ���'>����</a></li></ol>"&str
	set ajax=new AjaxXml
	ajax.re(split(str&"$$$","$$$"))
end sub

'��ȡ����Ϣ״̬
sub getpm()
	Dim rs,pmNumbers,ajax,username
	If not oblog.checkuserlogined() then
		pmNumbers=0
	Else
		Set rs=oblog.execute("select count(id) from oblog_pm where incept='"&oblog.l_uname&"' and isreaded=0 and delR=0")
		pmNumbers="("&rs(0)&")"
		set rs=nothing
	End If
	set ajax=new AjaxXml
	ajax.re(split(pmNumbers&"$$$","$$$"))
end sub



Sub savelog()
	if not oblog.checkuserlogined() then
		exit Sub
	Else
		If oblog.l_ulevel=6 Then
			Exit Sub
		End If
	end if
	Dim blog, logtext, i, rs, logid, isdraft, p, tid, log_tags, filename, log_files, log_Abstract
	Dim log_topic, log_text, log_face, log_time, log_classid, log_showword, log_blogteam, log_subjectid, log_password, log_ishide, log_istop, log_isencomment, log_isdraft, log_modiid, log_tb, log_filename, todraft, log_str, log_oldtb,log_teamsubject,log_isneedlogin,log_viewscores,log_viewgroupid
	Dim isblog, teamid,log_specialid,log_isTrouble
	dim restr,ajax
	set ajax=new AjaxXml
	isdraft = Int(Request("isdraft"))
	If oblog.l_Group(31,0) = 1 and isdraft<>1 Then
		If Not oblog.codepass Then
			oblog.adderrstr ("��֤�������ˢ�º��������룡�������ʾ����,��������ˢ��.")
			restr=split(oblog.errstr&"$$$0","$$$")
			ajax.re(restr)
			Response.End
		End if
	End If
	log_isTrouble=0
	logid=Request("logid")
	If logid<>"" Then logid=CLng(logid)
	log_oldtb = ""
	If logid=""  Then
		Dim sPostAccess
		sPostAccess=oblog.CheckPostAccess
		If sPostAccess<>"" Then
			oblog.ShowMsg sPostAccess,""
		End If
	End If

	log_topic = Replace_Plus(unescape(Trim(Request("topic"))))
	log_face = Request("face")
	If log_text = "" Then log_text = Replace_Plus(unescape(Trim(Request("edit"))))
	log_time = Request("selecty") & "-" & Request("selectm") & "-" & Request("selectd") & " " & Request("selecth") & ":" & Request("selectmi") & ":00"
	log_classid = Trim(Request("classid"))
	log_showword = Trim(Request("showword"))
	log_specialid= Trim(Request("specialid"))
	log_blogteam=CLng(Trim(Request("blogteam")))
	log_teamsubject=Trim(Request("blogteamsubject"))
	log_subjectid = Trim(Request("subjectid"))
	log_password = Trim(Request("ispassword"))
	log_isencomment = Trim(Request("isencomment"))
	log_ishide = Trim(Request("ishide"))
	log_istop = Trim(Request("istop"))
	log_tb = Trim(Request("tb"))
	log_filename = Trim(Request("filename"))
	log_isdraft = isdraft
	log_files = Trim(Request("log_files"))
	log_Abstract=Trim(Request("abstract"))
	log_isneedlogin=Trim(Request("isneedlogin"))
	log_viewscores=Trim(Request("viewscores"))
	log_viewgroupid=Trim(Request("viewgroupid"))
'	oblog.adderrstr (log_viewgroupid&"aa")
	If logid <>"" Then
		log_modiid = logid
	End If
	log_topic = Trim (log_topic)
	log_text = Trim (log_text)
	log_text = Replace(log_text, "#isubb#", "")
	If (log_topic = "" Or oblog.strLength(log_topic) > 120) and isdraft<>1 Then oblog.adderrstr ("��־���ⲻ��Ϊ��(���ܴ���120)��")
	if isdraft=1 and log_topic = "" then log_topic="����"
	If Trim(log_filename) = "�Զ����" Then log_filename = ""
	If (oblog.chkdomain(log_filename) = False And log_filename <> "") and isdraft<>1 Then oblog.adderrstr ("�ļ����Ʋ��Ϲ淶��ֻ��ʹ��Сд��ĸ�Լ����֣�")
	If log_text = "" Or oblog.strLength(log_text) > Int(oblog.Cacheconfig(34)) Then oblog.adderrstr (tName & "���ݲ���Ϊ���Ҳ��ܴ���" & oblog.Cacheconfig(34) & "�ַ���")
	Dim iChk1,iChk2,iChk3,iChk4
	if isdraft<>1 then
		iChk1=oblog.chk_badword(log_topic)
		iChk2=oblog.chk_badword(log_abstract)
		iChk3=oblog.chk_badword(log_text)
		iChk4=oblog.chk_badword(unescape(Trim(Request("logtags"))))
		If iChk1=0.1 Or iChk2=0.1 Or iChk3=0.1 Or iChk4=0.1 Then
			'��¼����һ��ϵͳ���ɲ��� ������һ��ֵ��ʱ��������û� OS*#046
			oblog.execute("Update oblog_user Set isTrouble=isTrouble+1 Where userid=" & oblog.l_uid)
			'дϵͳ��־
			Dim rstLog
			Set rstLog=Server.CreateObject("Adodb.Recordset")
			rstLog.Open "select * From oblog_syslog Where 1=0",conn,1,3
			rstLog.AddNew
			rstLog("username")=oblog.l_uname
			rstLog("addtime")=oblog.ServerDate(Now)
			rstLog("addip")=oblog.userip
			rstLog("desc")="�û�����"&oblog.l_uname & "(ID��" & oblog.l_uid & ")" & " �� " & oblog.ServerDate(Now()) & " �� " & oblog.userip & " ����һƪ���°������½�ֹ����Ĺؼ��֣����±���ֹ������:<br/><font color=red>��־���⣺" & EncodeJP(log_topic) & "<br/>���ɹؼ��֣�" & oblog.ShowBadWord & "</font>"
			rstLog("itype")=2 '�û���־��Դ
			rstLog.Update
			rstLog.Close
			oblog.adderrstr ("���ݻ��ǩ�д��ھ��Խ�ֹ�Ĺؼ���,��ע����������!")

			'�ж��Ƿ���Ҫ���
			If oblog.CacheConfig(13)<>"0" And  Trim(oblog.CacheConfig(13))<>"" Then
				Dim isRedirect
				rstLog.Open "select istrouble,lockuser From oblog_user Where userid=" & oblog.l_uid,conn,1,3
				If rstLog(0)>CInt(oblog.CacheConfig(13)) Then
					rstLog("lockuser")=1
					rstLog.Update
					rstLog.Close
					isRedirect = 1
					oblog.errstr= ""
					oblog.adderrstr ("�������������ֹ��࣬�Ѿ��������")
					'����û�(������ǿ���˳�) �û��� oblog.l_uName OS*#046

					Session ("CheckUserLogined_"&oblog.l_uName) = ""
					Oblog.CheckUserLogined()
				End If
			End If
			Set rstLog=Nothing
			If oblog.errstr <> "" Then
				If isRedirect = 1 Then
					restr=Split(Replace(oblog.errstr,"_","<br />")&"$$$3$$$index.asp","$$$")
				Else
					restr=Split(Replace(oblog.errstr,"_","<br />")&"$$$0","$$$")
				End If
				ajax.re(restr)
				Response.End
			End If
		Elseif iChk1 >=1 Or iChk2>=1 Or iChk3>=1 Then
			log_isTrouble=1
		End If
	end if
	If Not IsDate(log_time) Then oblog.adderrstr (tName & "ʱ���ʽ����")
	if log_teamsubject="" then log_teamsubject=0
	If log_showword = "" Then log_showword = 0
	If Not IsNumeric(log_showword) Then oblog.adderrstr (tName & "������ʾ��������Ϊ���֣�")
	If log_subjectid = "" Then log_subjectid = 0
	If log_classid = "" Then log_classid = 0
	If log_istop = "" Then log_istop = 0
	If log_isencomment = "" Then log_isencomment = 0
	If log_ishide = "" Then log_ishide = 0
	If log_isdraft = "" Then log_isdraft = 0
	'����־����ɲݸ壨�޸��ѷ�����־Ϊ�ݸ壩
	If Int(Request("oldisdraft")) = 0 And log_isdraft = 1 And log_modiid > 0 Then todraft = 1
	'���ݸ屣�����־���޸Ĳݸ�Ϊ����״̬��
	If Int(Request("oldisdraft")) = 1 And log_isdraft = 0 And log_modiid > 0 Then todraft = -1
	log_tags = Replace_Plus(unescape(Trim(Request.Form("logtags"))))
	If log_tags <> "" and isdraft<>1 Then
		log_tags = Replace(log_tags, "'", "")
		If Len(log_tags) > 255 Then
				oblog.adderrstr ("TAG�ܳ��Ȳ��ܴ���255���ַ�")
			End If
			If UBound(Split(log_tags, P_TAGS_SPLIT)) > (Int(oblog.CacheConfig(73)) - 1) Then
				oblog.adderrstr ("ÿƪ�������֧��" & oblog.CacheConfig(73) & "��TAG")
			End If
	End If
	If log_blogteam<>oblog.l_uId Then
		If CheckBlogTeam(log_blogteam) = False Then
			log_blogteam = oblog.l_uId
		End If
	End if
	set rs=Nothing

	If oblog.errstr <> "" Then
		restr=Split(Replace(oblog.errstr,"_","<br />")&"$$$0","$$$")
		ajax.re(restr)
		Response.end
	end If

    Set blog = New class_blog
    Set rs = Server.CreateObject("adodb.recordset")
    If log_modiid > 0 Then
        rs.open "select * from oblog_log where logid=" & log_modiid, conn, 2, 2
		'����һ������������û��޸���־��ʱ����־����Ϊ�ݸ壬���ύ������û��޸���ֱ���ύ���ֿ���
		If Int(Request("oldisdraft")) <> rs("isdraft") Then
			If log_isdraft = 0 Then
				todraft = -1
			End if
		End if
		If todraft = -1 Then
			Call oblog.GiveScore("",oblog.cacheScores(3),"")
			'��־��������
			rs("scores")=oblog.cacheScores(3)
		End if
    Else
        rs.open "select top 1 * from oblog_log Where 1=0 ", conn, 2, 2
        rs.addnew
		If log_isdraft = 0 Then
			Call oblog.GiveScore("",oblog.cacheScores(3),"")
			'��־��������
			rs("scores")=oblog.cacheScores(3)
		End if
    End If
    '��ʼд�����
    rs("topic") = EncodeJP(oblog.filt_astr(log_topic, 240))
    If Request("isubb") = "1" Then
        log_text = "#isubb#" & log_text
        rs("EditorType") = 1
    Else
        rs("EditorType") = 0
    End If
    log_text = EncodeJP(oblog.filtpath(oblog.filt_badword(log_text)))
    '���нű�����
    If oblog.l_Group(12,0)=0 Then log_text=FilterJS(log_text)
    '��һ������༭����ɵ�<DIV>&nbsp;</DIV>����,��ʹ��-1����
	log_text=Replace(log_text,"<DIV>&nbsp;</DIV>","<br/>")
	log_text=Replace(log_text,"<div>&nbsp;</div>","<br/>")
    rs("logtext") = log_text
    rs("face") = log_face
    rs("addtime") = log_time
    rs("classid") = log_classid
'	log_blogteam = oblog.l_uId
	if log_teamsubject>0 then log_subjectid=CLng(log_teamsubject)
	If rs("subjectid") <> Int(log_subjectid) And log_modiid > 0 Then
		oblog.Execute ("update oblog_subject set subjectlognum=subjectlognum+1 where subjectid=" & CLng (log_subjectid))
		oblog.Execute ("update oblog_subject set subjectlognum=subjectlognum-1 where subjectid=" & CLng (rs("subjectid")))
	End If
	rs("subjectid") = Int(log_subjectid)
	rs("showword") = Int(Trim(log_showword))
	If log_modiid = 0 Then
		rs("authorid") = oblog.l_uId
		rs("author") = EncodeJP(oblog.l_uName)
	End If
	rs("userid") = log_blogteam
	'--------------
	rs("is_log_default_hidden")=oblog.l_is_log_default_hidden '�Ƿ���ϵͳ��ҳ��ʾ����
	'--------------
	rs("ishide") = log_ishide
	rs("istop") = log_istop
	If log_modiid > 0 Then log_oldtb = rs("tburl")
	rs("tburl") = log_tb
	'�����ϴ��ļ�
	log_files=Replace(log_files," ","")
	'----------------------------------------^*^--
	log_files=FilterIds(log_files)

	If Left(log_files,1)="," Then log_files=Right(log_files,Len(log_files)-1)
	rs("logpics") = log_files
	rs("logtype") = 0
	rs("isencomment") = log_isencomment
	rs("Abstract") = log_Abstract
	rs("isneedlogin") = log_isneedlogin
	rs("viewscores") = log_viewscores
	rs("viewgroupid") = log_viewgroupid
	If log_ishide = 1 Or log_isneedlogin = 1 Or log_viewscores > 0 Or log_password <>"" Or log_viewgroupid > 0 Or oblog.l_blogpassword = 1 Then
		RS("IsSpecial") = 1
	Else
		RS("IsSpecial") = 0
	End If
	'��ѯ����־����ר���Ƿ�Ϊ���ص�
	Dim rssubject
	Set rsSubject = oblog.Execute ("SELECT ishide FROM oblog_subject WHERE subjectid = "&CLng (log_subjectid))
	If Not rsSubject.Eof Then
		If rsSubject(0) = 1 Then
			RS("IsSpecial") = OB_IIF(RS("IsSpecial"),0) + 1
		End If
	End if
	If rs("ispassword") = log_password Then

	Else
		If log_password <> "" Then
		   If log_password<>"�������룬�����޸��벻Ҫ����" Then rs("ispassword") = md5(Trim(log_password))
		Else
			log_password = ""
			rs("ispassword") = ""
		End If
	End If
	If oblog.l_Group(11,0) = 1 Then
		rs("passcheck") = 0
		log_Abstract = "����־��Ҫ����Ա��˺�ſɼ���"
	Else
		If logid = "" Then
			rs("passcheck") = 1
		End if
	End If
	If todraft <> 1 Then rs("isdraft") = log_isdraft
	rs("filename") = log_filename
	If log_specialid="" Then log_specialid=0
	rs("specialid") = log_specialid
	If log_modiid = 0 Then
		rs("iis") = 0
		rs("commentnum") = 0
		rs("trackbacknum") = 0
		rs("blog_password") = 0
		rs("truetime") = Now()
	End If
	rs("addip")=oblog.userip
	rs("istrouble")=log_isTrouble
	rs.Update
	rs.Close

	'---------------------------------------------------------------
	If (log_modiid = 0 And log_isdraft = 0) Or todraft = -1 Then
		Call OBLOG.log_count(log_blogteam,"",log_subjectid,log_classid,"+")
		oblog.Execute ("update [oblog_myurl] set isupdate=1 where mainuserid="&oblog.l_uid)
	End If
	If log_modiid = 0 Then
		Set rs = oblog.Execute("select max(logid) from oblog_log where userid=" & log_blogteam)
		tid = rs(0)
		rs.Close
	Else
		tid = log_modiid
	End If
	'�����ļ����� ##$
	If log_files <>"" Then
		oblog.Execute "Update oblog_upfile Set logid=" & tid & " Where fileid In (" & log_files & ")"
	End if
	'Tag����
	Call Tags_UserAdd(log_tags, oblog.l_uId, tid)
	If isdraft = 0 Then
		'�����״̬������ר���Ⱥ�鴦��
		If oblog.l_Group(11,0) = 0 or 1=1 Then
			'---------------------------------------------------------------
		    'ר�⴦��
			if log_specialid >0 Then
				log_specialid=CLng (log_specialid)
				rs.Open "select * From oblog_SpecialList Where logid=" & tid & " And specialid=" & log_specialid,conn,1,3
				If rs.Eof Then
					rs.Addnew
					oblog.Execute("Update oblog_Special Set s_count=s_count+1 Where specialid=" & log_specialid)
				End If
				rs("specialid")=log_specialid
				rs("userid")=oblog.l_uid
				rs("logid")=tid
				rs("author")=oblog.l_uname
				rs("topic")=log_topic
				rs("abstract")=log_abstract
				rs("addtime")=oblog.ServerDate(Now)
				rs("ispass")=0
				rs("istop")=0
				rs.Update
				rs.Close
			End If

			'Ⱥ�鴦��
			teamid=FilterIds(Request.Form("teamid"))
			If teamId<>"" Then
				teamId=Split(teamid,",")
				If Ubound(teamId) <=Int(oblog.CacheConfig(72)) Then
					For i=0 To Ubound(teamId)
						rs.Open "select * From oblog_teampost Where logid=" & tid & " And teamid=" & teamid(i),conn,1,3
						If rs.Eof Then
							rs.Addnew
							oblog.Execute "Update oblog_team Set icount1=icount1+1 Where teamId=" & teamid(i)
							rs("istop")=0
							rs("isbest")=0
							rs("ispass")=1
							rs("addtime")=log_time
							rs("addip")=oblog.userip
							rs("views")=0
							rs("replys")=0
							rs("scores")=1
							Call oblog.GiveScore("",oblog.cacheScores(13),"")
						End If
						rs("userid")=oblog.l_uid
						rs("author")=oblog.l_uname
						rs("teamid")=teamid(i)
						rs("logid")=tid
						rs("topic")=log_topic
						rs("content")=log_text
						rs("lastupdate")=oblog.ServerDate(Now)
						rs.Update
						rs.close
					Next
				End if
			End If
		End If
			'---------------------------------------------------------------
	    '������־��̬ҳ��
		blog.userid = log_blogteam
'		blog.isMulti=0
		blog.CreateFunctionPage
		blog.Update_log tid, 0
		'����ǹ�ͬ׫д��Ҫ������һƪ��־������
		If log_blogteam = oblog.l_uid Then
			If log_modiid = 0 Then
				set rs=oblog.execute("select top 1 logid from oblog_log where logid<"&tid&" and userid="&log_blogteam&" and logtype=0 order by logid desc")
				If Not rs.EOF Then blog.Update_log rs(0), 0
			End If
		End if
		blog.Update_calendar (tid)
		blog.Update_newblog (log_blogteam)
		blog.Update_Subject (log_blogteam)
		blog.Update_index 0
		blog.Update_info log_blogteam
		'����ǹ�ͬ׫д��������û���ҳ
		If log_blogteam <> oblog.l_uid Then
			blog.userid =  oblog.l_uid
			blog.CreateFunctionPage
			blog.Update_calendar (tid)
			blog.Update_newblog (oblog.l_uid)
			blog.Update_Subject (oblog.l_uid)
			blog.Update_index 0
			blog.Update_info oblog.l_uid
		End If
	    '��Ŀ�����ӷ���Pingָ��
	    If log_tb <> "" And log_tb <> log_oldtb Then
	        Dim objTrackBack,TrackBackIsOK
	        Set objTrackBack = New Class_TrackBack
			objTrackBack.logid = tid
	        objTrackBack.Blog_Name = blog.BlogName
	        objTrackBack.title = log_topic
	        objTrackBack.url = oblog.cacheConfig(3) & "go.asp?logid=" & tid
	        objTrackBack.Excerpt = log_topic & "<br />oBlog Created"
			TrackBackIsOK = objTrackBack.ProcessMultiPing(log_tb)
	        if TrackBackIsOK = True Then
				restr="������־�ɹ�������ͨ�淢���ɹ���$$$1"
			Else
				restr="������־�ɹ�������ͨ�淢��ʧ�ܣ�$$$1"
			End if
	        Set objTrackBack = Nothing
		Else
			restr="������־�ɹ�!$$$1"
			'������־�ɹ� ��־id  tid �û���  oblog.l_uName   OS*#046
	    End If
	Else
		If todraft = 1 Then
			logtodraft (tid)
		End If
		restr="����"&FormatDateTime(oblog.ServerDate(Now()),4)&"���浽�ݸ��䡣$$$2$$$"&tid
	End If
	'������־����������Session
	Session ("CheckUserLogined_"&oblog.l_uName) = ""
	Oblog.CheckUserLogined()
	Set rs = Nothing
	Set blog = Nothing
	ajax.re(Split(restr,"$$$"))
	Response.End()
End Sub

Sub logtodraft(logid)
    logid = CLng (logid)
    Dim uid, delname, subjectfile, fso, sid, rs,cid
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select userid,logfile,subjectfile,subjectid,isdraft,classid from oblog_log where logid=" & logid, conn, 1, 3
    If Not rs.EOF Then
		'����Ѿ��ǲݸ�����������
		If rs(4) = 1 Then rs.Close:Set rs = Nothing :Exit Sub
        uid = rs(0)
        delname = Trim(rs(1))
        subjectfile = rs(2)
        sid = rs(3)
		cid = rs(5)
        If delname <> "" Then
			If true_domain = 1 Then
				If InStr(delname, "archives") Then
					delname = Right(delname, Len(delname) - InStrRev(delname, "archives") + 1)
				Else
					delname = Right(delname, Len(delname) - InStrRev(delname, "/"))
				End If
				delname=oblog.l_udir&"/"&oblog.l_ufolder&"/"&delname
			End If
            Set fso = Server.CreateObject(oblog.CacheCompont(1))
            If fso.FileExists(Server.MapPath(delname)) Then fso.deleteFile Server.MapPath(delname)
        End If
        rs(1) = ""
        rs(4) = 1
        rs.Update
        rs.Close
		Call oblog.GiveScore("",-1*Abs(oblog.CacheScores(3)),"")
		Call OBLOG.log_count(uid,logid,sid,cid,"-")
        Dim blog
        Set blog = New class_blog
        blog.userid = uid
        'blog.update_index_subject 0,0,0,""
        blog.Update_index 0
        blog.Update_newblog (uid)
        Set blog = Nothing
        Set fso = Nothing
        Set rs = Nothing
    Else
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
End Sub

sub get_draft()
	dim rs,userid,draft_num,del_num,ajax
	userid=CLng(Request("userid"))
	set rs=oblog.execute("select count(logid) from oblog_log where isdraft=1 and isdel=0 and userid="&userid)
	draft_num=rs(0)
	if draft_num>0 then draft_num="("&draft_num&")" else draft_num=""
	set rs=oblog.execute("select count(logid) from oblog_log where isdel=1 and userid="&userid)
	del_num=rs(0)
	if del_num>0 then del_num="("&del_num&")" else del_num=""
	set rs=nothing
	set ajax=new AjaxXml
	ajax.re(split(draft_num&"$$$"&del_num,"$$$"))
end Sub
'�滻������־�е�"+"
Function Replace_Plus(str)
	Dim strTemp
	If str = "" Or IsNull(str) Then
		Replace_Plus= ""
		Exit Function
	End if
	strTemp=Replace (str,Chr(25),"+")
	Replace_Plus=strTemp
End Function
'����ŶӲ��ͷ�����־�ĺϷ���
Private Function CheckBlogTeam(tid)
	CheckBlogTeam = False
	Dim trs
	Set trs = oblog.Execute ("select id From oblog_blogteam WHERE otheruserid = " & oblog.l_uId & " AND mainuserid = " &tid)
	If Not trs.EOF Then CheckBlogTeam = True
	trs.close
	Set trs = Nothing
End Function
Sub digglog()
	If Not lcase(Request.ServerVariables("REQUEST_METHOD"))="post" Then Response.End
	Dim logid,ajax,restr,diggID,SQL,UID,authorid,username,diggNum,tstr,diggip,Pdigg
	Dim rsDigg,RSLog,RS,FromUrl
	logid = clng(Trim(Request("logid")))
	FromUrl = Trim(Request("fromurl"))
	diggip=oblog.UserIp
	If request("ip")<>"" Then diggip=CheckIP(request("ip"))

	On Error Resume Next
	response.clear
	set ajax=new AjaxXml
	If request("ptrue")=1 Then
		pdigg=oblog.checkuserlogined_digg(unescape(Trim(request("puser"))),Trim(request("ppass")))
		'ob_debug pdigg,1
		pdigg=Split(pdigg,"$$")

		If pdigg(0)=1 Then
			UID = pdigg(1)
			username = pdigg(2)
		Else
			UID = 0
			username = "(�ο�)"
		End If
	Else
		If oblog.checkuserlogined() Then
			UID = OBLOG.L_uid
			username = oblog.l_uname
		Else
			UID = 0
			username = "(�ο�)"
		End If
	End If
	If oblog.CacheConfig(83) = "0" And UID = 0 Then
		ajax.re(Split("<a href="&blogurl&"login.asp?fromurl="&Replace(FromUrl,"&","$")&">��¼</a>$$$1$$$"&logid&"$$$","$$$"))
		Exit Sub
	End If
	if not IsObject(conn) then link_database
	Set RSLog = Server.CreateObject("adodb.recordset")
	RSLog.open "SELECT authorid,classid,Abstract,logfile,topic,author,diggnum,logtext FROM oblog_log WHERE isdel=0 and (isspecial=0 or isspecial is null) and logid = "&CLng (logid),CONN,1,3
	If RSLog.EOF Then
		restr = "ʧ��$$$1$$$"&logid&"$$$"
	Else
		authorid = RSLog(0)
		'�����Լ����Լ��ƵĲ���
		If UID=authorid Then
			restr = "����$$$1$$$"&logid&"$$$"
			ajax.re(Split(restr,"$$$"))
			Exit Sub
		End If
		Set rsDigg = Server.CreateObject("adodb.recordset")
		rsDigg.open "SELECT * FROM oblog_userdigg WHERE logid = "&logid&" AND authorid="&authorid,conn,1,3
		If rsDigg.EOF Then
			rsDigg.AddNew
			rsDigg("diggtitle") = RSLog(4)
			rsDigg("diggurl") = RSLog(3)
			rsDigg("diggnum") = 0
			rsDigg("diggdes") = OB_IIF(RSLog(2),Left(RemoveHtml(RSLog(7)),255))
			rsDigg("authorid") = authorid
			rsDigg("classid") = RSLog(1)
			rsDigg("logid") = logid
			rsDigg("author") = RSLog(5)
			rsDigg("addip") = diggip
			rsDigg("istate")  = 1
			rsDigg.Update
			rsDigg.movelast
			tstr = rsDigg.BookMark
		Else
			If rsDigg("istate") = 0 Then ajax.re(Split("$$$1$$$"&logid&"$$$�ܾ�","$$$")): Exit Sub
			diggID = rsDigg("diggID")
			diggNum = OB_IIF(rsDigg("diggNum"),0)
		End If
		If IsEmpty(diggID) Then
			rsDigg.BookMark = tstr
			diggID = rsDigg("diggID")
			diggNum = 0
		End If
		Set rs = Server.CreateObject("adodb.recordset")
		If UID > 0 Then
			SQL ="SELECT * FROM oblog_digg WHERE userid = "&UID&" AND diggid="&diggID
		Else
			SQL ="SELECT * FROM oblog_digg WHERE addip = '"&diggip&"' AND diggid="&diggID
		End If
		rs.open SQL,CONN,1,3
		If Not rs.EOF Then
			restr = "����$$$1$$$"&logid&"$$$"&diggNum
		Else
			rs.AddNew
			rs("userid") = UID
			rs("diggid") = diggID
			rs("addip") = diggip
			rs("logid") = logid
			rs("authorid") = authorid
			rs("username") = username
			If UID = 0 Then rs("isguest") = 1
			rs.Update
			diggNum = diggNum + 1
			RSLog("diggnum") = diggNum
			RSLog.Update
			restr = "�ɹ�$$$2$$$"&logid&"$$$"&diggNum
			oblog.Execute ("UPDATE oblog_userdigg SET diggnum = "&diggNum&",lastdiggtime="&G_Sql_Now&" WHERE diggID = "&diggID)
			oblog.Execute ("UPDATE oblog_user SET diggs = diggs + 1  WHERE userid = "&authorid)
			'�ӷֲ���
			Call oblog.GiveScore("",oblog.cacheScores(22),authorid)
		End If
		diggID = Empty
	End If
	Set rs = Nothing
	Set rsDigg = Nothing
	Set RSLog = Nothing
	ajax.re(Split(restr,"$$$"))
	Response.End
End Sub
Sub SaveReport()
	Dim RS,userid,authorid,ajax,username
	Dim report_type,logid
	report_type = Request("report_type")
	logid = CLng(Trim(Request("logid")))
	If report_type <>"" Then report_type = CLng(report_type) Else report_type = 0
	set ajax=new AjaxXml
	If oblog.checkuserlogined() Then
		userid = OBLOG.L_uid
		username = oblog.l_uname
	Else
		userid = 0
		username = "�ο�"
	End if
	Set RS = oblog.Execute ("SELECT authorid FROM oblog_log WHERE logid = "&logid)
	If RS.EOF Then
		ajax.re(Split("��־������$$$1$$$"&logid,"$$$"))
		Response.End
	Else
		authorid = RS(0)
		Set RS = Nothing
		if not IsObject(conn) then link_database
		Set RS = Server.CreateObject("ADODB.RecordSet")
		RS.open "SELECT * FROM oblog_digg WHERE 1 = 0",CONN,1,3
		RS.AddNew
		rs("userid") = userid
		rs("diggid") = 0
		rs("addip") = oblog.UserIp
		rs("logid") = logid
		rs("authorid") = authorid
		rs("diggtype") = report_type
		rs("username") = username
		If userid = 0 Then rs("isguest") = 1
		rs.Update
	End If
	ajax.re(Split("��л���ķ��������ǻἰʱ����$$$2$$$"&logid,"$$$"))
	Response.End
End Sub
%>