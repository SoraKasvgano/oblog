<%
'Oblog Group Class
'Class_Group.asp
'teamusers��state��״̬
'teamusers: state 1��Ч;2�������3������4 ������Ա 5 ����Ա
'�ܾ���ɾ���ü�¼,����ͨ������ɾ����ϵͳ���Զ���һ����Ϣ���û�
Class Class_Team
	Public g_id,g_Name,g_Ico,CssFile,g_Links,g_Creater,g_ManagerId,g_ManagerName,g_CreateTime,g_ViewLimit
	Public g_Placard,icount0,icount1,icount2,g_intro,g_ViewPassWord,g_Domain,g_DomainRoot,g_OtherPost
	Public g_announce,g_guide,team_Domain,g_URL
	Public PageFrameWork,PageBody,ShowMode,ErrMsg
	Private iPage,Sql,rs,imMode,pid,icoGood,icoTop,icoBlog,groupPWD,fileID
	Private g_Show_main,g_Show_log,g_show_title,g_Show_list

	Private Sub Class_initialize()
		Set rs=Server.CreateObject("Adodb.RecordSet")
		iPage=12
		On Error Resume Next
        If Not IsObject(conn) Then Link_DataBase
        pid=1
	End Sub

	Private Sub Class_terminate()
		On Error Resume Next
        If IsObject(conn) Then conn.Close: Set conn = Nothing
		If ErrMsg<>"" Then Response.Write ErrMsg
    End Sub

	Public Property Let GroupId(byval Value)
		g_id=Int(Value)
		rs.Open "select * From oblog_team Where teamid=" & g_id,conn,1,1
		If rs.Eof Then
			Response.Write "Ŀ��" &oblog.CacheConfig(69)& "������!"
			Response.End
		Else
			If rs("iState") = 1 Then
				Response.Write "Ŀ��" &oblog.CacheConfig(69)& "��δ������Ա���!"
				Response.End
			ElseIf rs("iState") = 2 Then
				Response.Write "Ŀ��" &oblog.CacheConfig(69)& "������!"
				Response.End
			End if
		End If
		rs.Filter="iState=3"

		If Not Rs.EOF Then
			g_Name=rs("t_name")
			g_Ico=rs("t_ico")
			g_CreateTime=rs("CreateTime")
			g_ManagerId=rs("ManagerId")
			g_ManagerName=rs("ManagerName")
			g_ViewLimit = rs("ViewLimit")
			g_ViewPassWord = rs("ViewPassWord")
			g_Domain = rs("t_domain")
			g_DomainRoot = rs("t_domainroot")
			g_OtherPost = rs("otherpost")
			icount0 = rs("icount0")
			icount1 = rs("icount1")
			icount2 = rs("icount2")
			g_intro = rs("intro")
			'��ȡ������Ϣ
			g_announce=OB_IIF(rs("announce"),"û������")
			'��ȡ����������Ϣ
			g_links=OB_IIF(rs("links"),"û������")
			'������Ȩ��
			If g_ViewLimit = 0 Then
				If Not IsMember Then
					response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" /><title>�������</title><style></style>"
					Response.Write "��" &oblog.CacheConfig(69)& "Ϊ˽��" &oblog.CacheConfig(69)& ",��" &oblog.CacheConfig(69)& "��Ա��Ȩ���ʴ�" &oblog.CacheConfig(69)&"<br/>�����Ը�" &oblog.CacheConfig(70)&"����,��������" &oblog.CacheConfig(69)& "."

					Response.End
				End If
				ErrMsg = ""
			ElseIf g_ViewLimit = 1 Then
				If Not IsNull(g_ViewPassWord) And g_ViewPassWord<>"" Then
					groupPWD = Request.Cookies(cookies_name)("group_pwd_"&g_id)
					If groupPWD = "" Or groupPWD<> g_ViewPassWord Then
						Response.Redirect blogurl&"chkblogPassword.asp?groupid="&g_id&"&fromurl="&Replace(oblog.GetUrl,"&","$")
						Response.End
					End If
				End if
			End if
			Call GetTheme
			Call IsManager
			If oblog.CacheConfig(5)="1" And oblog.CacheConfig(75) <> "" And g_Domain<>"" And Not IsNull(g_Domain) Then
				team_Domain = "http://"&g_Domain&"."&g_DomainRoot&"/"
			Else
				team_Domain = oblog.cacheConfig(3)&"group.asp?gid="&g_id
			End If
			g_URL = vbcrlf & "<!-- " &oblog.CacheConfig(69)& "��ַ -->" & vbcrlf & "<div id=""GroupUrl""><a href="""&team_domain&""">"&team_domain&"</a></div>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "��ַ end -->"& vbcrlf
			g_guide = vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� -->" & vbcrlf & "<ul id=""GroupMenu"">" & vbcrlf & "	<li><a href=""group.asp?gid=$group_id$"">"&oblog.CacheConfig(69)&"��ҳ</a></li>"&vbcrlf&"	<li><a href=""group.asp?cmd=list&gid=$group_id$"">"&oblog.CacheConfig(69)&"����</a></li>"&vbcrlf&"	<li><a href=""group.asp?cmd=good&gid=$group_id$"">"&oblog.CacheConfig(69)&"����</a></li>"&vbcrlf&"	<li><a href=""group.asp?cmd=users&gid=$group_id$"">��Ա�б�</a></li>"&vbcrlf&"	<li>$group_m_buttons$<a href=""group.asp?cmd=join&gid=$group_id$"">�������</a></li>"&vbcrlf&"	<li><a href=""group.asp?cmd=album&gid=$group_id$"">������</a></li>"&vbcrlf&"	<li><a href=""group.asp?cmd=post&gid=$group_id$"">��������</a></li>"&vbcrlf&"	<li><a href=""group.asp?cmd=postphoto&gid=$group_id$"">������Ƭ</a></li>"&vbcrlf&"</ul>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� end -->"&vbcrlf
		Else
			Response.Write "Ŀ��" &oblog.CacheConfig(69)& "�Ѿ���ɾ��!"
			Response.End
		End If
		rs.Close
		Set rs=Nothing
	End Property

	Public Sub Show()
		PageFrameWork=MakeMainPage(0)
		iMode=Request("mode")
		select Case ShowMode
			Case 1
				PageBody=GetIndexList(iMode,False)
			Case 2
				PageBody=ShowPost(pid)
			Case 3
				PageBody=GetUser(g_id)
			Case Else
				PageBody=ErrMsg
		End select
		'If ShowMode Then
		PageFrameWork=Replace(PageFrameWork,"$title$",g_show_title)
		PageFrameWork=Replace(PageFrameWork,"$group_list$",PageBody)
		Response.Write PageFrameWork
		PageFrameWork=""
	End Sub

	'��ʾ��һ��־���ظ�
	Public Sub ShowPost(pid)
		Dim sRet,sPost,sReply,sEditor
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$title$",g_Name&"--������Ӽ��ظ�")
		sRet=Replace(sRet,"$group_list$",GetPost(pid))
		Response.Write sRet
		sRet=""
	End Sub

	'��ʾ��־�б�
	Public Sub ShowList(iType)
		Dim sRet
		sRet=MakeMainPage(iType)
		If iType = 0 Then
			sRet=Replace(sRet,"$group_posts$",GetIndexList(iType,True))
		Else
			sRet=Replace(sRet,"$title$",g_Name&"--��������б�")
			sRet = Replace(sRet,"$group_list$",vbcrlf & "<div id=""GroupList"">" & vbcrlf & "	<div class=""title"">��־�б�</div>" & vbcrlf & "$group_list$" & vbcrlf & "	<div class=""clear""></div>" & vbcrlf)
			If iType = -1 Then
				sRet = Replace(sRet,"$group_list$",GetIndexList(0,False))
			Else
				sRet = Replace(sRet,"$group_list$",GetIndexList(iType,False))
			End if
		End if
		Response.Write sRet
		sRet=""
	End Sub

	Public Sub ShowUsers()
		Dim sRet
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$title$",g_Name&"--��Ա�б�")
		sRet = Replace(sRet,"$group_list$",vbcrlf & "<div id=""GroupList"">" & vbcrlf & "	<div class=""title"">��Ա�б�</div>" & vbcrlf & "$group_list$" & vbcrlf & "	<div class=""clear""></div>" & vbcrlf & "</div>")
		sRet=Replace(sRet,"$group_list$",GetUsers)
		Response.Write sRet
		sRet=""
	End Sub

	Public Sub ShowLinksForm()
		Dim sRet
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$group_list$",LinksForm)
		Response.Write sRet
		sRet=""
	End Sub

	Public Sub ShowPlacardForm()
		Dim sRet
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$group_list$",PlacardForm)
		Response.Write sRet
		sRet=""
	End Sub

	Public Sub PostForm()
		Dim sRet
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$title$",g_Name&"--��������")
		sRet = Replace(sRet,"$group_list$",vbcrlf & "$group_list$")
		sRet=Replace(sRet,"$group_list$",CommentForm(postid,0))
		Response.Write sRet
		sRet=""
	End Sub


	Public Sub ShowJoinForm()
		Dim sRet
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$title$",g_Name&"--�������")
		sRet = Replace(sRet,"$group_list$",vbcrlf & "<div id=""GroupList"">" & vbcrlf & "	<div class=""title"">�������</div>$group_list$" & vbcrlf & "	<div class=""clear""></div>" & vbcrlf & "</div>")
		sRet=Replace(sRet,"$group_list$",JoinForm(g_id))
		Response.Write sRet
		sRet=""
	End Sub

	Public Sub ActionJoin()
		Dim sRet
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$title$",g_Name&"--ͨ������")
		sRet=Replace(sRet,"$group_list$",AcceptJoin())
		Response.Write sRet
		sRet=""
	End Sub

	Public Sub PostPHOTO()
		Dim sRet
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$title$",g_Name&"--������Ƭ")
		sRet = Replace(sRet,"$group_list$",vbcrlf & "<div id=""GroupList"">" & vbcrlf & "	<div class=""title"">������Ƭ</div>" & vbcrlf & "$group_list$" & vbcrlf & "	<div class=""clear""></div>" & vbcrlf & "</div>")
		sRet=Replace(sRet,"$group_list$","<iframe id='d_file' frameborder='0' src='upload.asp?re=no&isphoto=1&tMode=2&teamid="&g_id&"' width='320' height='400' scrolling='no'></iframe>")
		Response.Write sRet
		sRet=""
	End Sub

	Public Sub album()
		Dim sRet
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$title$",g_Name&"--������")
		sRet = Replace(sRet,"$group_list$",vbcrlf & "<div id=""GroupList"">" & vbcrlf & "	<div class=""title"">������</div>" & vbcrlf & "$group_list$" & vbcrlf & "	<div class=""clear""></div>" & vbcrlf & "</div>")
		sRet=Replace(sRet,"$group_list$",ShowPhoto())
		Response.Write sRet
		sRet=""
	End Sub

	Public Sub photocomment()
		Dim sRet,sPhoto
		GetPhotoComment sPhoto
		sRet=MakeMainPage(1)
		sRet=Replace(sRet,"$title$",g_Name&"--"&g_show_title)
		sRet=Replace(sRet,"$group_list$",sPhoto)
		Response.Write sRet
		sRet=""
	End Sub

	'iMode��1�Ƽ� 2 �������� 3 �ǲ������� ���� ȫ��
	'isIndex �Ƿ�Ϊ��ҳ����
	Function GetIndexList(iMode,isIndex)
		If Not isIndex Then
			icoBlog="<img src=""oBlogStyle/group/01.gif"" border=""0""  title=""��ͨ����"" />"
			icoGood="<img src=""oBlogStyle/group/02.gif""  border=""0"" title=""��������"" />"
			icoTop="<img src=""oBlogStyle/group/03.gif"" border=""0""  title=""�ö�����"" />"
	'		icoBlog="[��ͨ����]"
	'		icoGood="[��������]"
	'		icoTop="[�ö�����]"
		End if
		Dim SqlPart,sRet,sRet1,i,r,Nums
		Dim rs,lPage,lAll,lPages,sTitle,sMBar
		select Case iMode
			Case "1"
				SqlPart=" And isbest=1 "
			Case "2"
				SqlPart=" And isblog=1 "
			Case "3"
				SqlPart=" And isblog=0 "
			Case Else
		End select
		G_P_Filename="group.asp?cmd="&cmd&"&gid=" & g_id & "&mode="&imode&"&page="
		sRet=""
		If isIndex Then Nums = 8 Else Nums = 500
		Set rs=Server.CreateObject("Adodb.RecordSet")
		If isIndex Then
			Sql="select  * from (select top "&Nums&" isbest,istop,logid,postid,topic,author,replys,lastupdate,addtime,views From oblog_teampost Where teamid=" & g_id & " And iDepth=0 " & SqlPart & " Order By Lastupdate Desc) AS T"
		Else
			Sql="select  * from (select top "&Nums&" isbest,istop,logid,postid,topic,author,replys,lastupdate,addtime,views From oblog_teampost Where teamid=" & g_id & " And iDepth=0 " & SqlPart & " Order By istop DESC ,Lastupdate Desc) AS T"
		End If
'		Sql= Sql & " union "
'		Sql= Sql & " (select  top "&Nums&" isbest,istop,logid,postid,topic,author,replys,lastupdate,addtime,views From oblog_teampost Where teamid=" & g_id & " And iDepth=0 And isTop=0  " & SqlPart & " Order by Lastupdate desc )) DERIVEDTBL ORDER BY istop DESC,Lastupdate DESC"
		rs.Open Sql,conn,1,1
'		Response.Write(sql)
		'Set rs=oblog.Execute(Sql)
		If rs.Eof Then
			If iMode = "1" Then
				rs.Close
				sRet= sRet & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "�������� -->" & vbcrlf
				sRet= sRet & "<div id=""GroupBestLog"">" & vbcrlf
				sRet= sRet & "Ŀǰ��û���κ�����" & vbcrlf
				sRet= sRet & "</div>" & vbcrlf
				sRet= sRet & "<!-- " &oblog.CacheConfig(69)& "�������� end -->" & vbcrlf
				GetIndexList=sRet
				sRet=""
				Exit Function
			Else
				rs.Close
				sRet= sRet & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "�������� -->" & vbcrlf
				sRet= sRet & "<div id=""GroupNewLog"">" & vbcrlf
				sRet= sRet & "Ŀǰ��û���κ�����" & vbcrlf
				sRet= sRet & "</div>" & vbcrlf
				sRet= sRet & "<!-- " &oblog.CacheConfig(69)& "�������� end -->" & vbcrlf
				GetIndexList=sRet
				sRet=""
				Exit Function
			End if
		End If
		'��ҳ
		If Request("page") = "" Or Request("page") ="0" then
			lPage = 1
		Else
			lPage = Int(Request("page"))
		End If
		lAll=Int(rs.recordcount)
		'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
		rs.CacheSize = iPage
		rs.PageSize = iPage
		rs.movefirst
		lPages = rs.PageCount
		If lPage>lPages Then lPage=lPages
		rs.AbsolutePage = lPage
		Do While Not rs.Eof And i < rs.PageSize
		'д����
			sTitle=""
			If rs("istop")=1 Then
				sTitle= icoTop  & sTitle
			ElseIf rs("isbest")=1 Then
				sTitle= icoGood  & sTitle
			Else
				If rs("logid")>0 Or 1=1 Then sTitle= icoBlog  & sTitle
			End if
			if Int(i/2)*2=i then r=1 else r=2
			If isIndex Then
				sRet1="				<td class=""t1"">"&sTitle&"<a href=""group.asp?cmd=show&gid="& g_Id &"&pid=" & rs("postid")&""" title="""&OB_IIF(rs("topic"),"����")&""">"&OB_IIF(rs("topic"),"����")&"</a></td>" & vbcrlf
				sRet1=sRet1&"				<td class=""t2"">"&rs("replys")&"</td>" & vbcrlf
				sRet1=sRet1&"				<td class=""t3""><a href=""go.asp?user="&rs("author")&""" title="""&rs("author")&""">"&rs("author")&"</a><span>"&rs("Lastupdate")&"</span></td>" & vbcrlf
	'			sRet1=sRet1&"<td class='s4'>"&rs("Lastupdate")&"</td>"
				sRet1= vbcrlf & "			<tr class=""r"&r&""">" & vbcrlf &sRet1&"			</tr>" & vbcrlf
			Else
				sRet1="			<td class=""t1"">"&sTitle&"</td>" & vbcrlf
				sRet1=sRet1&"			<td class=""t2""><a href=""group.asp?cmd=show&gid="& g_Id &"&pid=" & rs("postid")&""" title="""&OB_IIF(rs("topic"),"����")&""">"&OB_IIF(rs("topic"),"����")&"</a></td>" & vbcrlf
				sRet1=sRet1&"			<td class=""t3""><a href=""go.asp?user="&rs("author")&""" title="""&rs("author")&""">"&rs("author")&"</a><span>"&rs("addtime")&"</span></td>" & vbcrlf
				sRet1=sRet1&"			<td class=""t4"">"&rs("replys")&"/"&rs("views")&"</td>" & vbcrlf
				sRet1=sRet1&"			<td class=""t5""><span>"&rs("Lastupdate")&"</span></td>" & vbcrlf
				sRet1= vbcrlf & "		<tr class=""r"&r&""">" & vbcrlf &sRet1&"		</tr>" & vbcrlf
			End if
			sRet=sRet & sRet1 & vbcrlf
			i=i+1
			rs.MoveNext
		Loop
		rs.Close
		Set rs=Nothing
		'����һ���ײ���ҳ��
		If isIndex Then
			If iMode = "1" Then
				sRet= vbcrlf & "<!-- " &oblog.CacheConfig(69)& "�������� -->" & vbcrlf & "<div id=""GroupBestLog"">" & vbcrlf & "	<table id=""LogList"">" & vbcrlf & "		<thead>" & vbcrlf & "			<tr>" & vbcrlf & "				<th class=""t1"">����</th>" & vbcrlf & "				<th class=""t2"">�ظ�</th>" & vbcrlf & "				<th class=""t3"">���ߣ�������</th>" & vbcrlf & "			</tr>" & vbcrlf & "		</thead>" & vbcrlf & "		<tbody>" & vbcrlf & sRet & vbcrlf & "		</tbody>" & vbcrlf & "	</table>" & vbcrlf & "</div>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "�������� end -->" & vbcrlf
			Else
				sRet= vbcrlf & "<!-- " &oblog.CacheConfig(69)& "�������� -->" & vbcrlf & "<div id=""GroupNewLog"">" & vbcrlf & "	<table id=""LogList"">" & vbcrlf & "		<thead>" & vbcrlf & "			<tr>" & vbcrlf & "				<th class=""t1"">����</th>" & vbcrlf & "				<th class=""t2"">�ظ�</th>" & vbcrlf & "				<th class=""t3"">���ߣ�������</th>" & vbcrlf & "			</tr>" & vbcrlf & "		</thead>" & vbcrlf & "		<tbody>" & vbcrlf & sRet & vbcrlf & "		</tbody>" & vbcrlf & "	</table>" & vbcrlf & "</div>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "�������� end -->" & vbcrlf
			End if
		Else
			sRet="<table id=""GroupLogList"">" & vbcrlf & "	<thead>" & vbcrlf & "		<tr>" & vbcrlf & "			<th class=""t1""></th>" & vbcrlf & "			<th class=""t2"">����</th>" & vbcrlf & "			<th class=""t3"">����</th>" & vbcrlf & "			<th class=""t4"">�ظ������</th>" & vbcrlf & "			<th class=""t5"">������</th>" & vbcrlf & "		</tr>" & vbcrlf & "	</thead>" & vbcrlf & "	<tbody>"&sRet& vbcrlf & "	</tbody>" & vbcrlf & "</table>"
		End if
		If Not isIndex Then sRet=sRet & "<div id=""GroupPages"">" & vbcrlf & PageBarNum(lAll,iPage,lPage,G_P_Filename) & vbcrlf & "</div>"
		'���ݽű�����
		'sRet= sRet & vbcrlf & "<div id=""comment_list""></div>"
		GetIndexList=sRet
		sRet=""

	End Function

	Function GetUsers()
		Dim sRet
		Dim rs,lPage,lAll,lPages,i
		G_P_Filename="group.asp?gid=" & g_id & "&cmd="&cmd&"&page="
		Sql="select a.userid,a.province,a.city,username,nickname,blogname,user_icon1,log_count,user_group,scores From oblog_user a,"
		Sql= Sql & "(select  userid,state From oblog_teamusers Where Teamid=" & g_id & ") b Where a.userid=b.userid and b.state>2 Order By b.state Desc"
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open sql,conn,1,1
		If rs.EOF Then
			GetUsers="<div> ����Ա�ʺŲ����ڻ����Ѿ���ɾ�� </div>"
			Exit Function
		End if
		If Request("page") = "" Or Request("page") ="0" then
			lPage = 1
		Else
			lPage = Int(Request("page"))
		End If
		lAll=Int(rs.recordcount)
		'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
		rs.CacheSize = iPage
		rs.PageSize = iPage
		rs.movefirst
		lPages = rs.PageCount
		If lPage>lPages Then lPage=lPages
		rs.AbsolutePage = lPage
		Do While Not rs.Eof and i < rs.PageSize
			sRet= sRet & "		<ul class=""UserList"">" & vbcrlf
			sRet= sRet & "			<li class=""userimg""><a href=""go.asp?userid=" & rs("userid") & """ target=""_blank""><img src=""" & OB_IIF(rs("user_icon1"),""&blogurl&"images/ico_default.gif") & """ border=0 /></a></li>" & vbcrlf
			sRet= sRet & "			<li class=""username""><a href=""go.asp?userid=" & rs("userid") & """ target=""_blank"">" & rs("username") & "</a></li>" & vbcrlf
			sRet= sRet & "			<li class=""usercity"">(" & rs("province") & rs("city")  &")</li>" & vbcrlf
			sRet= sRet & "		</ul>" & vbcrlf
			i=i+1
			rs.Movenext
		Loop
		rs.Close
		Set rs=Nothing

		GetUsers="	<div id=""GroupBestUser"">" & vbcrlf & sRet & vbcrlf & "	</div>"
		sRet=""
		GetUsers=GetUsers & "<div id=""GroupPages"">" & vbcrlf & PageBarNum(lAll,iPage,lPage,G_P_Filename) & vbcrlf & "</div>"
	End Function


	Sub SaveComment()
		Dim title,content,author,userid,url,sql,rs,pid,iDepth,modify
		modify=Trim(Request("modify"))
		author=Request.Form("username")
		pid=Request("pid")
		If pid="" Then
			pid=0
			iDepth=0
		Else
			iDepth=1
		End If
		pid=CLng (pid)
		title=RemoveHtml(Request.Form("commenttopic"))
		content=Request.Form("oblog_edittext")
		'��֤��У��
		if oblog.CacheConfig(30)=1 Then
			If  Request("CodeStr")="" then
				oblog.ShowMsg "��֤������뷵��ˢ�º��������룡",""
				exit sub
			Else
				if not oblog.codepass then
					oblog.ShowMsg "��֤������뷵��ˢ�º��������룡",""
					exit sub
				end if
			End If
		end if
		If Len(content)=0 Or Len(content)>50000 Then
			oblog.ShowMsg "���������ݲ���Ϊ�գ��ҳ��Ȳ��ܴ���50000",""
			exit sub
		End If
		If oblog.checkuserlogined() Then
			Author= oblog.l_uName
			userid=oblog.l_uid
		End If
		If Len(Author)=0 Or Len(Author)>20 Then
			oblog.ShowMsg "�û�������Ϊ�գ��ҳ��Ȳ��ܴ���20",""
			exit sub
		End If
		If IsMember=False Then
			If pid=0 Then
				oblog.ShowMsg "�Ǳ�" &oblog.CacheConfig(69)& "��Ա�����Է������⣬���ɻظ�����������������" &oblog.CacheConfig(69)& "",""
				exit Sub
			Else
				If g_OtherPost = 0  And Not g_ViewLimit="-1" Then
					oblog.ShowMsg "�Ǳ�" &oblog.CacheConfig(69)& "��Ա��Ȩ����ظ�����������������" &oblog.CacheConfig(69)& "",""
					exit Sub
				End if
			End If
		End If
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.Open "select * From oblog_teampost Where postid=" & pid,conn,1,3
		If rs.Eof Then
			If pid>0 Then
				rs.Close
				Set rs=Nothing
				ErrMsg= "Ŀ�����ⲻ����"
				Exit Sub
			End If
		else
			If IsManager=False And modify="1" then
				if rs("userid")<>oblog.l_uid then
					rs.Close
					Set rs=Nothing
					ErrMsg= "��Ȩ��"
					Exit Sub
				end if
			end if
		End If
		If pid > 0 And modify <>"1" Then
			title = "Re:"&rs("topic")
		End if
		if modify<>"1" then
			rs.AddNew
			rs("teamid")=g_Id
			rs("author")=Author
			rs("parentid")=pid
			rs("iDepth")=iDepth
			rs("logid")=0
			rs("userid")=OB_IIF(userid,0)
			rs("addip")=oblog.userip
			rs("addtime")=oblog.ServerDate(Now)
			rs("LastUpdate")=oblog.ServerDate(Now)
			rs("ispass")=1
			rs("istop")=0
			rs("isbest")=0
		end if
		if pid>0 and modify<>"1" then
			rs("content")=oblog.Ubb_comment(EncodeJP(oblog.filt_badword(content)))
		else
			rs("content")=EncodeJP(oblog.filt_badword(content))
		end if
		rs("topic")=EncodeJP(oblog.InterceptStr(oblog.filt_badword(title),50))
		rs.Update
		if modify<>"1" Then
			If pid>0 Then
				oblog.Execute("Update oblog_teampost Set replys=replys+1,LastUpdate='" & oblog.ServerDate(Now) & "' Where teamid=" & g_id &" and postid="&pid)
				oblog.Execute("Update oblog_team Set icount2=icount2+1 Where teamid=" & g_id)
			Else
				oblog.Execute("Update oblog_team Set icount1=icount1+1 Where teamid=" & g_id)
			End If
		End if
		If userid>0 and modify<>"1" Then oblog.Execute("Update oblog_teamusers Set post_replys=post_replys+1 Where userid=" & userid & " And teamid=" & g_id)
		rs.Close
		Set rs=Nothing
		If pid=0 Then
			Response.Redirect "group.asp?gid=" & g_id
		Else
			Response.Redirect "group.asp?gid=" & g_id & "&pid=" & pid
		End If
	End Sub

	'����ģʽ���ö�/ȡ���ö�/����/ȡ������/ɾ��
	Sub PostManage(cmd,pid)
		Dim targetUrl
		If IsManager=false and cmd<>"del" Then
			ErrMsg= "��û��Ȩ�޽��д˲���"
			Exit Sub
		End If
		pid=CLng (pid)
		select Case Cstr(cmd)
			Case "good1"
				Sql="Update oblog_teampost Set isbest=1 Where postid=" & pid
				targetUrl= "group.asp?gid=" & g_id & "&pid=" &pid
			Case "good0"
				Sql="Update oblog_teampost Set isbest=0 Where postid=" & pid
				targetUrl= "group.asp?gid=" & g_id & "&pid=" &pid
			Case "top1"
				Sql="Update oblog_teampost Set istop=1 Where postid=" & pid
				targetUrl= "group.asp?gid=" & g_id & "&pid=" &pid
			Case "top0"
				Sql="Update oblog_teampost Set istop=0 Where postid=" & pid
				targetUrl= "group.asp?gid=" & g_id & "&pid=" &pid
			Case "del"
				if IsManager then
					sql="select userid,parentid From oblog_teampost Where postid=" & Pid
				else
					if oblog.CheckUserLogined then
						sql="select userid,parentid From oblog_teampost Where postid=" & Pid&" and userid="&oblog.l_uid
					else
						exit sub
					end if
				end if
				Set rs=oblog.Execute(sql)
				If Not rs.Eof Then
					oblog.Execute "Update oblog_teampost Set replys=replys-1,scores=scores-1 Where postid=" & rs(1)
					'Response.End()
					If Ob_IIF(rs(0),0)>0 Then oblog.Execute "Update oblog_teamusers Set post_all=post_all-1 ,post_replys=post_replys-1 Where userid=" & rs(0) & " And teamid=" & g_Id
					oblog.Execute "Delete From  oblog_teampost Where postid=" & pid
					Sql="Update oblog_team Set icount1=icount1-1 Where teamid=" & g_id
				end if
				targetUrl= "group.asp?gid=" & g_id
				'oblog.Execute Sql
			Case "6"
				Dim rs
				Set rs=oblog.Execute("select userid,parentid From oblog_teampost Where postid=" & Pid)
				If Not rs.Eof Then
					oblog.Execute "Delete From  oblog_teampost Where postid=" & pid
					'���½���ظ���Ŀ
					oblog.Execute "Update oblog_teampost Set replys=replys-1,scores=scores-1 Where postid=" & rs(1)
					If Ob_IIF(rs(0),0)>0 Then oblog.Execute "Update oblog_teamusers Set post_all=post_all-1 ,post_replys=post_replys-1 Where userid=" & rs(0) & " And teamid=" & g_Id
					Sql="Update oblog_team Set icount2=icount2-1 Where teamid=" & g_id
				End If
				'ɾ��һ���ظ�
				Sql="Delete From  oblog_teampost Where postid=" & pid
				oblog.Execute Sql
				targetUrl= "group.asp?gid=" & g_id & "&pid=" &pid
		End select
		If Sql<>"" Then oblog.Execute Sql
		oblog.ShowMsg "�����ɹ�!",targetUrl
		'Response.Redirect targetUrl
	End Sub
	Function Search()

	End Function

	'-----------------------------
	'������ʾҳ��
	'-----------------------------

	'��������ҳ��ģ�� action 0 ��ģ�� 1 ��ģ��
	Private Function MakeMainPage(action)
		Dim sMList,sMButtons,sRet,sRet1,rs,Sql
		'����ģ�崦��
		If action = 0 Then
			sRet=g_Show_main
		Else
			sRet=g_Show_list
		End if
		If postid = 0 And fileID = 0 Then
			If action = 0 Then sRet=Replace(sRet,"$title$",g_Name)
		End If
		sRet=Replace(sRet,"$group_ico$",ProIco(g_Ico,2))
		sRet=Replace(sRet,"$group_url$",g_URL)
'		sRet=Replace(sRet,"$title$",g_show_title)
		sRet=Replace(sRet,"$group_guide$",g_guide)
		sRet=Replace(sRet,"$group_m_buttons$","")
		sRet=Replace(sRet,"$group_id$",g_Id)
		sRet=Replace(sRet,"$group_name$",vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� -->" & vbcrlf & "<div id=""GroupName"">"&g_Name&"</div>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� end -->")
		If OB_IIF(g_Ico,"")="" Then
			sRet=Replace(sRet,"$group_ico$",g_Ico)
		Else
			sRet=Replace(sRet,"$group_ico$","")
		End If
		sRet=Replace(sRet,"$group_creater$",g_creater)
		sRet=Replace(sRet,"$group_id$",g_Id)
		'�ײ�
		sRet=Replace(sRet,"$group_bottom$",oblog.CacheConfig(10))
		sRet=Replace(sRet,"$group_comments$",getminilist())
		'����
		sRet=Replace(sRet,"$group_placard$",vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� -->" & vbcrlf & "<div id=""GroupPlacard"">" & vbcrlf & g_announce & vbcrlf & "</div>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� end -->" & vbcrlf)
		'��������
		sRet=Replace(sRet,"$group_links$",vbcrlf & "<!-- " &oblog.CacheConfig(69)& "�������� -->" & vbcrlf & "<div id=""GroupLinks"">" & vbcrlf & g_links & vbcrlf & "	<div id=""ad_teamlinks""></div>" & vbcrlf & "</div>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "�������� end -->" & vbcrlf)
		'Ⱥ����Ϣ
		sRet=Replace(sRet,"$group_info$",GetGroupInfo)
		'��Ծ�û�
		sRet=Replace(sRet,"$group_bestuser$",index_showuser("best"))
		'���¼���
		sRet=Replace(sRet,"$group_newuser$",index_showuser("new"))
		sRet=Replace(sRet,"$group_admin$",GetAdminList)
		sRet=Replace(sRet,"$group_bestposts$",GetIndexList(1,True))
		sRet=Replace(sRet,"$group_photo$",showPhoto())
		sRet = sRet &vbcrlf&"<span id=""ad_teambot""></span></body>"&vbcrlf
		sRet = sRet &"</html>"
		sRet = sRet & "<script src=""" & blogurl&"ShowXml.asp?teamid="&g_id&"""></script>"
		'����Ⱥ��
		'���ԾȺ��
		MakeMainPage=sRet
	End Function

	function index_showuser(action)
		dim sql,tmp,i
		if action="best" then
			tmp="Order By post_replys Desc"
		else
			tmp="Order By addtime Desc"
		end if
		Sql="select a.*,b.nickname,b.username,b.user_icon1 From (select top 9 * From oblog_teamusers Where state>2 and teamid=" & g_Id & " "&tmp&") a,oblog_user b Where a.userid=b.userid"
		'Response.Write(sql)
		Set rs=oblog.Execute(Sql)
		If Not rs.Eof Then
				If action = "best" Then
					index_showuser=index_showuser & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "��Ծ��Ա -->" & vbcrlf & "<div id=""GroupBestUser"">" & vbcrlf
				Else
					index_showuser=index_showuser & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���³�Ա -->" & vbcrlf & "<div id=""GroupNewUser"">" & vbcrlf
				End If
			Do While Not rs.Eof
				i=i+1
				if i>9 then exit Do
				index_showuser=index_showuser & "	<ul class=""UserList"">" & vbcrlf
				index_showuser=index_showuser & "		<li class=""userimg""><a href="""&blogurl&"go.asp?userid=" & rs("userid")&""" target=_blank><img src=""" & OB_IIF(rs("user_icon1"),""&blogurl&"images/ico_default.gif")&""" /></a></li>" & vbcrlf
				index_showuser=index_showuser & "		<li class=""username""><a href=""go.asp?userid=" & rs("userid")&"""  target=_blank>" & OB_IIF(rs("nickname"),rs("username"))&"</a></li>" & vbcrlf
				index_showuser=index_showuser & "	</ul>" & vbcrlf
				rs.movenext
			Loop
					index_showuser=index_showuser & "</div>" & vbcrlf
				If action = "best" Then
					index_showuser=index_showuser & "<!-- " &oblog.CacheConfig(69)& "��Ծ��Ա end -->" & vbcrlf
				Else
					index_showuser=index_showuser & "<!-- " &oblog.CacheConfig(69)& "���³�Ա end -->" & vbcrlf
				End If

		End If
		Set rs=Nothing
	end function

	'���������ö���������ҳ
	Function GetPost(byval pid)
		Dim i,rs,Sql,sTitle,sMBar
		Dim lPage,lAll,lPages
		Dim sRet,sRet1
		Set rs=Server.CreateObject("Adodb.RecordSet")
		Sql="select a.*,b.user_icon1 From oblog_teampost a,oblog_user b Where a.postid="& pid &" and a.userid=b.userid And iDepth=0"
		G_P_Filename="group.asp?cmd=show&gid="&g_Id&"&pid=" & pid & "&page="
		Set rs=oblog.execute(Sql)
		If rs.Eof Then
			sRet="<li>���Ϊ" & pid & "�����ⲻ����</li>"
			Set rs=Nothing
			Exit Function
		End If
		Oblog.Execute("UPDATE oblog_teampost SET views = views + 1 WHERE postid =  "&pid)
		sTitle=OB_IIF(rs("topic"),"����")
		If rs("isbest")=1 Then
			sMBar="<a href=""group.asp?cmd=good0&gid=" & g_Id & "&pid=" & rs("postid") & """>ȡ������</a> | "
		Else
			sMBar="<a href=""group.asp?cmd=good1&gid=" & g_Id & "&pid=" & rs("postid") & """>��Ϊ����</a> | "
		End If
		If rs("istop")=1 Then
			sMBar=" | <a href=""group.asp?cmd=top0&gid=" & g_Id & "&pid=" & rs("postid") & """>ȡ���ö�</a> | " & sMBar
		Else
			sMBar=" | <a href=""group.asp?cmd=top1&gid=" & g_Id & "&pid=" & rs("postid") & """>��Ϊ�ö�</a> | " & sMBar
		End If
		'If rs("logid")>0 Then sTitle= icoBlog  & sTitle
		sMBar=sMBar & "<a href=""group.asp?cmd=del&gid=" & g_Id & "&pid=" & rs("postid") & """ onclick=""return confirm('ȷ��ɾ�������ӣ�');"">ɾ��</a> | "
		sMBar=sMBar & "<a href=""group.asp?cmd=post&modify=1&gid=" & g_Id & "&pid=" & rs("postid") & """>�༭</a>"
		g_show_title = sTitle
		sRet = g_Show_log
		sRet = vbcrlf & "		<div class=""LogList"">"& vbcrlf &sRet
		sRet = sRet &"		</div>" & vbcrlf
		'�滻ID,������ʾ
		sRet=Replace(sRet,"$group_name$",vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� -->" & vbcrlf & "<div id=""GroupName"">"&g_Name&"</div>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� end -->")
		sRet=Replace(sRet,"c_content_down", "c_content_down1")
		sRet=Replace(sRet,"$group_post_title$", "<div class=""LogTitle"">"&sTitle&" (���������"&rs("views")&")</div>")
		sRet=Replace(sRet,"$group_content$",rs("content"))
		sRet=Replace(sRet,"$group_post_userico$",ProIco(rs("user_icon1"),2))
		sRet=Replace(sRet,"$group_post_user$",rs("author"))
		sRet=Replace(sRet,"$group_post_time$",rs("addtime"))
		sRet=Replace(sRet,"$group_post_content$",filtscript(rs("content")))
		sRet=Replace(sRet,"$group_post_id$",rs("postid"))
		sRet=Replace(sRet,"$group_post_replys$","<a href=""#add_comment"">�ظ�("&rs("replys")&")</a> ")
		sRet=Replace(sRet,"$group_tags$",OB_IIF(rs("tags"),""))
		sRet=Replace(sRet,"$group_post_link$","#")
		sRet=Replace(sRet,"$group_post_userurl$","go.asp?user="&rs("author"))
		sRet=Replace(sRet,"$group_post_high$","¥��")

		If imMode=1 Then
			sRet=Replace(sRet,"$group_post_m$",sMBar)
		Else
			if oblog.CodeCookie(rs("author"))=Request.Cookies(cookies_name)("username") then
				sRet=Replace(sRet,"$group_post_m$","<a href=""group.asp?cmd=del&gid=" & g_Id & "&pid=" & rs("postid") & """>ɾ��</a> | <a href=""group.asp?cmd=post&modify=1&gid=" & g_Id & "&pid=" & rs("postid") & """>�༭</a>")
			else
				sRet=Replace(sRet,"$group_post_m$","")
			end if
		End If
		rs.Close
		'�ҹ��
		sRet=sRet & vbcrlf & "<div id=""oblog_ad_team_post_1""></div>" & Vbcrlf
		'������
		rs.Open "Select a.*,b.User_Icon1,b.Username From (select top 500 * From oblog_teampost Where idepth>0 And parentid=" & pid & " Order By postid Desc) a Left Join oblog_user b  On a.userid =b.userid order by a.postid asc",conn,1,1
		If rs.Eof Then
			sRet1=""
		Else
			i=0
			'��ҳ
			If Request("page") = "" Or Request("page") ="0" then
				lPage = 1
			Else
				lPage = Int(Request("page"))
			End If
			lAll=Int(rs.recordcount)
			'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
			rs.CacheSize = iPage
			rs.PageSize = iPage
			rs.movefirst
			lPages = rs.PageCount
			If lPage>lPages Then lPage=lPages
			rs.AbsolutePage = lPage
			sRet1 = ""
			i=0
			Do While Not rs.Eof And i < rs.PageSize
				i=i+1
				sRet1=sRet1 & vbcrlf & Replace("<div class=""CommentsList""><a name=""a_"&rs("postid")&""" />" &g_Show_log&"</div>","$group_topic$","") & vbcrlf
				sRet1=Replace(sRet1,"$group_post_title$", "<div class=""CommentsTitle"">Re:"&sTitle&"</div>")
				sRet1=Replace(sRet1,"$group_post_userurl$","go.asp?user="&rs("author"))
				sRet1=Replace(sRet1,"$group_post_replys$","")
				sRet1=Replace(sRet1,"$group_content$","<span id=""c_"&rs("postid")&""">"&rs("content")&"</span>")
				sRet1=Replace(sRet1,"$group_post_user$","<a href='go.asp?user="&rs("author")&"'><span id=""n_"&rs("postid")&""">"&rs("author")&"</span></a>")
				sRet1=Replace(sRet1,"$group_post_time$","<span id=""t_"&rs("postid")&""">"&rs("addtime")&"</span>")
				sRet1=Replace(sRet1,"$group_post_userico$",OB_IIF(rs("user_icon1"),"images/ico_default.gif"))
				sRet1=Replace(sRet1,"$group_post_high$","��<span class=""xx"">" & i & "</span>¥")
				If imMode=1 or oblog.CodeCookie(rs("author"))=Request.Cookies(cookies_name)("username") Then
					sRet1=Replace(sRet1,"$group_post_m$","<a href=""javascript:reply_quote('"& rs("postid")&"')"" >����</a> | <a href=""group.asp?cmd=del&gid=" & g_id &"&pid=" & rs("postid")& """>ɾ��</a>")
				Else
					sRet1=Replace(sRet1,"$group_post_m$","<a href=""javascript:reply_quote('"& rs("postid")&"')"" >����</a> ")
				End If
				rs.MoveNext
			Loop
		End If
		sRet1= vbcrlf &  "<div id=""comment_list"">" & sRet1 &"</div>" & vbcrlf
		'�ҷ�ҳ����
		sRet1= sRet1 & "<div id=""GroupPages"">" & vbcrlf & PageBarNum(lAll,iPage,lPage,G_P_Filename) & vbcrlf & "</div>"
		'�һظ�
		sRet1=sRet1
		GetPost = sRet&sRet1
		GetPost = vbcrlf & "<div id=""GroupList"">" & vbcrlf & "	<div class=""title"">"&g_Name&" &gt; �����б� </div>" & vbcrlf & "	<div id=""Log_List"">"&GetPost
		GetPost = GetPost & vbcrlf & "</div>" & vbcrlf & "</div>" & vbcrlf
		GetPost = GetPost & CommentForm(pid,0)
	End Function

	'��ȡ���»ظ�
	Function GetNewComments()
		Dim rs,Sql,sRet
		Sql="select top 5 content From oblog_teampost Where iDepth>0 And teamid=" & g_id & " Order By postid Desc"
		Set rs=oblog.Execute(Sql)
		If rs.Eof Then
			sRet="<li>-<li>"
		Else
			Do While Not rs.Eof
				sRet=sRet & "<li>" & Left(RemoveHtml(rs(0)),10) & "...</li>"
				rs.MoveNext
			Loop
		End if
		Set rs=Nothing
		GetNewComments=sRet
		sRet=""
	End Function

	'���Ⱥ����Ϣ
	Function GetGroupInfo()
		Dim sRet
		sRet=""
		sRet=sRet & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���� -->" & vbcrlf
		sRet=sRet & "<div id=""GroupInfo"">" & vbcrlf
		sRet=sRet & "	<ul class=""Groupico"">" & vbcrlf
		sRet=sRet & "		<li><img class=""group_img"" src="""&ProIco(g_ico,2)&""" onload=""rsimg(this,195);"" /></li>" & vbcrlf
		sRet=sRet & "		<li><span>" &oblog.CacheConfig(69)& "���ƣ�</span>"&g_Name&"</li>" & vbcrlf
		sRet=sRet & "	</ul>" & vbcrlf
		sRet = sRet & "	<div class=""GroupIntro""><span>"&oblog.CacheConfig(69)&"���ܣ�</span><p>"&g_intro&"</p></div>" & vbcrlf
		sRet=sRet & "	<ul class=""GroupData"">" & vbcrlf
		sRet=sRet & "		<li><span>�����ߣ�" & g_ManagerName & "</span></li>" & vbcrlf
		sRet=sRet & "		<li><span>����ʱ�䣺" & g_createtime & "</span></li>" & vbcrlf
		sRet=sRet & "		<li><span>��Ա������" & icount0 & "</span></li>" & vbcrlf
		sRet=sRet & "		<li><span>����������" & icount1 & "</span></li>" & vbcrlf
		sRet=sRet & "		<li><span>�ظ�������" & icount2 & "</span></li>" & vbcrlf
		sRet=sRet & "	</ul>"& vbcrlf
		sRet=sRet & "</div>"& vbcrlf
		sRet=sRet & "<!-- " &oblog.CacheConfig(69)& "���� end -->" & vbcrlf
		GetGroupInfo = sRet
	End Function


	function GetAdminList()
		Dim rs,rst,sRet
		Sql="select TOP 4 a.userid,a.province,a.city,username,nickname,blogname,user_icon1,log_count,user_group,scores From oblog_user a,"
		Sql= Sql & "(select  userid,state,addtime From oblog_teamusers Where Teamid=" & G_id & ") b Where a.userid=b.userid and b.state=5 Order By b.addtime Desc"
		Set rs=Oblog.Execute(Sql)
		Do While Not rs.Eof
			sRet= sRet & "	<ul class=""GroupAdmin"">	<li class=""Adminimg""><img src=""" & OB_IIF(rs("user_icon1"),""&blogurl&"images/ico_default.gif") & """ border=0 width=48 height=48></li>" & vbcrlf
			sRet= sRet & "		<li class=""Adminname""><a href=""go.asp?userid=" & rs("userid") & """ target=""_blank"">" &  rs("username") &"</a></li>" & vbcrlf
			sRet= sRet & "		<li class=""Admincity"">(" & rs("province") & rs("city")  &")</li>" & vbcrlf
			sRet= sRet &"	</ul>"
			rs.MoveNext
		Loop
		GetAdminList= vbcrlf & "<!-- " &oblog.CacheConfig(69)& "����Ա -->" & vbcrlf & "<div id=""GroupAdmin"">" & vbcrlf & "	" & vbcrlf &sRet& "" & vbcrlf & "</div>" & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "����Ա end -->" & vbcrlf
	end  function

	'��ȡ��ص�Ⱥ���б���Ϣ
	Function GetTeams(byval sNumber,byval sType)
		Dim Sql,rs,sRet,sField1,sField2
		select Case sType
			Case 1
				' hot ����,�ظ����
				sField1="icount2"
				sField2="icount2 Desc"
			Case 2
				'active �����,�����������
				sField1="icount1"
				sField2="icount1 Desc"
			Case 3
				'���Ӵ�
				sField1="icount0"
				sField2="icount0 Desc"
			Case 4
				'���¼���
				sField= "icount0"
				sField= "teamid Desc"
		End select
		Sql="select top " & sNumber & " teamid,t_name, " &  sField & " From oblog_team Order by " & sField
		Set rs=oblog.Execute(Sql)
		If rs.Eof Then
			sRet="<li>��û���κ�" & oblog.CacheConfig(69) &"��Ϣ</li>"
		Else
			Do While Not rs.Eof
				sRet=sRet & "<li><a href=""group.asp?Group_id=" & rs(0) & """ target=""_blank"">" & rs(1) & "</a>(" & rs(2) & ")</li>"
				rs.Movenext
			Loop
		End If
		Set rs=Nothing
		GetHotTeams=sRet
		sRet=""
	End Function

	Function CommentForm(id,action)
		If IsMember=False Then
			If pid=0 Then
				oblog.ShowMsg "�Ǳ�" &oblog.CacheConfig(69)& "��Ա�����Է������⣬���ɻظ�����������������" &oblog.CacheConfig(69)& "","group.asp?cmd=join&gid="&g_id
				exit Function
			Else
				If g_OtherPost = 0 And Not g_ViewLimit="-1" Then
					oblog.ShowMsg "�Ǳ�" &oblog.CacheConfig(69)& "��Ա��Ȩ����ظ�����������������" &oblog.CacheConfig(69)& "","group.asp?cmd=join&gid="&g_id
					exit Function
				End if
			End If
		End If
		Dim sName,sRet,sTopic,sContent,modify,sql
		Dim FormUrl
		If action = 0 Then
			FormUrl = "group.asp?cmd=save&gid="&g_Id&"&pid=" & id &"&modify="&Trim(Request("modify"))
		ElseIf action = 1 Then
			FormUrl = "SaveAlbumComment.asp?fileid="&id&"&teamid="&g_id
		End if
		modify=Trim(Request("modify"))
		If oblog.checkuserlogined()=false Then
			CommentForm="<p><a href='login.asp?fromurl=group.asp?cmd="&cmd&"$gid="&g_Id&"$pid="&id&"'>�������¼����ܽ��лظ����߷����µ�����</a></p>"
			Exit Function
		End If
		if Trim(Request("modify"))="1" and id<>"" then
			if IsManager=true then
				sql="select * from oblog_teampost where postid="&CLng(id)
			Else
				If IsMember Then
					sql="select * from oblog_teampost where postid="&CLng(id)&" and userid="&oblog.l_uid
				End if
			end if
			set rs=oblog.execute(sql)
			if not rs.eof then
				sTopic=rs("topic")
				sContent=rs("content")
			end if
		end If
		sRet="<a name=""add_comment""></a>" & vbcrlf & "<div id=""form_comment"">" & vbcrlf
		If id <> "" And modify<>"1" Then
			sRet = sRet & "	<div class=""title"">�ظ�����</div>" & vbcrlf
		End if
		sRet = sRet & "<form action='"&FormUrl&"' method='post' name='commentform' id='commentform'>" & vbcrlf
		sRet=sRet&"<div id=""ad_teamcomment""></div>" & vbcrlf
		sName=oblog.l_uname
		If  sName ="" Then  sName="�ο�"
		sRet=sRet & "	<fieldset>" & vbcrlf
		If cmd="post" Then
		sRet=sRet & "		<legend>" & sName & " , ��ӭ����" &oblog.CacheConfig(69)& "����,�ڴ˴����������ݽ�����ʾ�����Ĳ�����</legend>" & vbcrlf
		else
		sRet=sRet & "		<legend>" & sName & " , ��ӭ������" &oblog.CacheConfig(69)& "�ظ�,�ڴ˴����������ݽ�����ʾ�����Ĳ�����</legend>" & vbcrlf
		End If
		sRet=sRet & "		<table>" & vbcrlf
		If Not oblog.checkuserlogined() Then
			sRet=sRet & "			<tr >" & vbcrlf
			sRet=sRet & "				<td><label for=""UserName"">�û�����</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""UserName"" type=""text"" id=""UserName"" size=""15"" maxlength=""20"" value="""" /></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
		Else
			sRet=sRet & "			<tr style=""display:none"">" & vbcrlf
			sRet=sRet & "				<td><label for=""UserName"">�ǳƣ�</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""UserName"" type=""text"" id=""UserName"" size=""15"" maxlength=""20"" value="""" /></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
			sRet=sRet & "			<tr style=""display:none"">" & vbcrlf
			sRet=sRet & "				<td><label for=""Password"">���룺</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name='Password' type='password' id='Password' size='15' maxlength='20' value='' />&nbsp;(�ο�������������)</td>" & vbcrlf
			sRet=sRet & "			</tr style=""display:none"">" & vbcrlf
			sRet=sRet & "			<tr style=""display:none"">" & vbcrlf
			sRet=sRet & "				<td><label for=""homepage"">��ҳ��</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""homepage"" type=""text"" id=""homepage"" size=""42"" maxlength=""50"" value=""http://"" /></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
		End If
		If Id="" or modify="1" Then
			sRet=sRet & "			<tr>" & vbcrlf & "				<td><label for=""commenttopic"">���⣺</label></td>" & vbcrlf & "				<td><input name=""commenttopic"" type=""text"" id=""commenttopic"" size=""50"" maxlength='50' value="""&sTopic&""" /></td>" & vbcrlf & "			</tr>" & vbcrlf
			sRet=sRet & "			<tr>" & vbcrlf & "				<td><label>���ݣ�</label></td>" & vbcrlf & "				<td><div id=""oblog_edit""><span id=""loadedit"" style=""font-size:12px""><img src='"&blogurl&"images/loading.gif' align='absbottom'> ��������༭��...</span><textarea id=""oblog_edittext"" name=""oblog_edittext"" style=""width:400px;height:250px; display:none"" >"&sContent&"</textarea></div></td>" & vbcrlf & "			</tr>" & vbcrlf
		Else
			sRet=sRet & "			<tr>" & vbcrlf & "				<td><label>���ݣ�</label></td>" & vbcrlf & "				<td><div id=""oblog_edit""><img src="""&blogurl&"images/loading.gif""></div><textarea id=""oblog_edittext"" name=""oblog_edittext1"" style=""width:400px;height:250px; display:none"" >"&sContent&"</textarea></div></td>" & vbcrlf & "			</tr>" & vbcrlf
		end If
		if oblog.CacheConfig(30)="1" Then
			sRet=sRet & "			<tr id =ob_code>" & vbcrlf
			sRet=sRet & "				<td><label for=""CodeStr"">��֤�룺</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""CodeStr"" id=""CodeStr"" type=""text"" size=""6"" maxlength=""20"" /> "&oblog.getcode&"</td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
		End if
		sRet=sRet & "			<tr>" & vbcrlf
		sRet=sRet & "				<td></td>" & vbcrlf
		sRet=sRet & "				<td><input type=""submit"" id=""Submit"" value="" �� �� ""  /></td>" & vbcrlf
		sRet=sRet & "			</tr>" & vbcrlf
		sRet=sRet & "		</table>"& vbcrlf
		sRet=sRet & "	</fieldset>"& vbcrlf
		sRet=sRet & "</form>" & vbcrlf
		sRet=sRet & "</div>" & vbcrlf

		if oblog.CacheConfig(30)="1" Then
			sRet=sRet & "<script>document.getElementById(""ob_code"").style.display='';</script>"
		end if
		if id<>"" and modify<>"1" then
			sRet=sRet & "<script>function addcode(){return true;}</script>"
			sRet=sRet & "<script src=""commentedit.asp""></script>"
		else

		end If
		if id="" or modify="1" then
			'����༭��
			sRet=sRet&	"<script language=JavaScript src='"&C_Editor_UBB&"/scripts/language/schi/editor_lang.js'></script>"
			sRet=sRet&	"<script language=JavaScript src='"&C_Editor_UBB&"/scripts/innovaeditor.js'></script>"
			sRet=sRet&	"<script language=""JavaScript"">"
			sRet=sRet&	"var oEdit1 = new InnovaEditor(""oEdit1"");"
			sRet=sRet&	"oEdit1.width=397;"
			sRet=sRet&	"oEdit1.height=260;"
			sRet=sRet&	"oEdit1.features=[""Hyperlink"",""Image"",""Flash"",""Media"",""CustomObject"",""|"",	""ClearAll"",""PasteWord"",""PasteText"",""RemoveFormat"",""|"",	""Bold"",""Italic"",""Underline"",""Strikethrough"",""|"",							""ForeColor"",""BackColor"",""|""];"
			sRet=sRet&	"oEdit1.cmdCustomObject = ""modelessDialogShow('"&blogdir&"editor/scripts/emot.htm',280,200)""; "
			sRet=sRet&	"oEdit1.cmdAssetManager=""modalDialogShow('"&blogdir&"editupload.asp',640,465)"";"
			sRet=sRet&	"oEdit1.REPLACE(""oblog_edittext"");"
			sRet=sRet&	"oEdit1.focus();"
			sRet=sRet&	"</script>"
			'�༭���������
		End if
		CommentForm=sRet
		sRet=""
	End Function

	Function GetTheme()
		Dim sRet,sStyle,oFso,oStream
		Dim team_Show
		Dim trs
		team_Show = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"&vbcrlf
		team_Show = team_Show &"<html xmlns=""http://www.w3.org/1999/xhtml"">"&vbcrlf
		team_Show = team_Show &"<head>"&vbcrlf
		team_Show = team_Show &"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"&vbcrlf
		team_Show = team_Show &"<title>$title$</title>"&vbcrlf
		team_Show = team_Show &"<script src="""&blogurl&"inc/main.js"" type=""text/javascript""></script>"&vbcrlf
		team_Show = team_Show &"{OB_STYLE}"&vbcrlf
		team_Show = team_Show &"</head>"&vbcrlf
		team_Show = team_Show &"<body>"&vbcrlf
		team_Show = team_Show &"<span id=""ad_teamtop""></span>"&vbcrlf
		g_Show_main = team_Show
'		If Application(oblog.cache_name&"_group_theme_main"&g_id)="" Then
			Set trs = oblog.Execute ("select user_skin_main,user_skin_showlog FROM oblog_team WHERE teamid = "&g_id)
			If IsNull(trs(0)) Or IsNull(trs(1)) Then
				set trs=oblog.Execute("select skinmain,skinshowlog from oblog_teamskin where isdefault=1")
				If trs.EOF Then
					set trs=oblog.Execute("select TOP 1 skinmain,skinshowlog from oblog_teamskin")
				End  if
			End if
			sRet = trs(0)
			sStyle = OB_PickUpCss(sRet)
			g_Show_main = Replace(g_Show_main,"{OB_STYLE}",sStyle)
			g_Show_main = g_Show_main & sRet
			g_Show_list = team_Show&trs(1)
			g_Show_list = Replace(g_Show_list,"{OB_STYLE}",sStyle)
			trs.Close
			Set trs = Nothing
'			Application.Lock
			'ģ��
			Set oFso=Server.CreateObject(oblog.CacheCompont(1))
			Set oStream=oFSO.OpenTextFile(Server.Mappath("oBlogStyle/group/g_log.htm"),1,False)
			g_Show_log = oStream.ReadAll
'			Application(oblog.cache_name&"_group_theme_list"&g_id) = g_Show_log
'			Application(oblog.cache_name&"_group_theme_main"&g_id) = g_Show_main
'			Application(oblog.cache_name&"_group_theme_post"&g_id) = g_Show_list
'			Application.Unlock
			sRet=""
'		Else
'			g_Show_main = Application(oblog.cache_name&"_group_theme_main"&g_id)
'			g_Show_list  = Application(oblog.cache_name&"_group_theme_post"&g_id)
'			g_Show_log = Application(oblog.cache_name&"_group_theme_list"&g_id)
'		End If
	End Function

	Function IsManager()
		Dim userin,Min,isMin,sql
		isMin=False
		IsManager=false
		imMode=0
		userin= ProtectSQL(oblog.filt_badstr(Request.Cookies(cookies_name)("username")))
		sql="SELECT top 1 userid FROM oblog_teamusers WHERE (state = 5) AND (teamid =  " & G_id & ") AND (userid = (SELECT TOP 1 userid   FROM oblog_user WHERE (username ='"&userin&"')))"
		Set Min= Server.CreateObject("adodb.recordset")
		Min.open sql, conn, 1, 1
		If Not (Min.eof Or Min.bof) Then
		If Min(0)<>"" And Not IsNull(Min(0)) Then isMin=True
		End If
		Min.close
		Set Min=Nothing
		If isMin Then
			If oblog.checkuserlogined()=true Then
				imMode=1
				IsManager=true
			End If
		End If

	End Function

	Function IsMember()
		Dim rs
		IsMember=false
		If oblog.checkuserlogined()=true Then
			Set rs=oblog.Execute("select id From oblog_teamusers Where state>2 and teamid=" & g_id & " And userid=" & oblog.l_uid )
			If Not rs.Eof Then
				IsMember=true
			End If
			Set rs=Nothing
		End If
	End Function
	'----------------------------------------------------
	'Ⱥ�������ģ��
	'----------------------------------------------------
	'�������ģ��
	Function JoinForm(id)
		Dim sRet,rs
'		If oblog.checkuserlogined()=false Then
'			JoinForm="<p><a href='login.asp?fromurl=group.asp?cmd=join$gid="&g_id&"'>�������ȵ�¼������������</a></p>"
'			Exit Function
'		End If
		'�жϼ�������
		Set rs=oblog.execute("select joinlimit,joinscores,icount0 From oblog_team Where teamid="& CLng (id))
		If rs.Eof Then
			ErrMsg="Ŀ��" &oblog.CacheConfig(69)& "������!"
			Response.End
		End If
		select Case rs(0)
			Case 1
			Case 2
				ErrMsg="��" &oblog.CacheConfig(69)& "ֻ����"&oblog.CacheConfig(70)&"�������룬�����������"
			Case 3
				If oblog.l_uscores<rs(1) Then
					ErrMsg="���뱾" &oblog.CacheConfig(69)& "��Ҫ���� " & rs(1) & " �����,���Ļ��ֲ���"
				End If
		End select
		if rs(2)>=Int(oblog.CacheConfig(71)) then
			ErrMsg="��" &oblog.CacheConfig(69)& "��Ա�Ѵﵽϵͳ����"&oblog.CacheConfig(71)&"�ˡ�"
		end if
		Set rs=Nothing
		If ErrMsg<>"" Then
			JoinForm=ErrMsg
			Exit Function
		End If
		'�Ƿ��κ��˶����Լ���
		sRet="<form id=""join"" action=group.asp?cmd=savejoin&gid="&g_id&" method=""post"">" & vbcrlf
		sRet = sRet &"	<fieldset>" & vbcrlf
		If Not oblog.checkuserlogined() Then
			sRet=sRet & "		<legend>�ο�,����д����������Ϣ</legend>" & vbcrlf
			sRet=sRet & "		<table>" & vbcrlf
			sRet=sRet & "			<tr>" & vbcrlf
			sRet=sRet & "				<td><label for=""UserName"">�û�����</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""UserName"" type=""text"" id=""UserName"" size=""15"" maxlength=""20"" value="""" /></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
			sRet=sRet & "			<tr>" & vbcrlf
			sRet=sRet & "				<td><label for=""Password"">���룺</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""Password"" type=""password"" id=""Password"" size=""15"" maxlength=""20"" value="""" /></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
		Else
			sRet=sRet &"		<legend>"& oblog.DecodeCookie(Request.Cookies(cookies_name)("username")) & ",����д����������Ϣ</legend>"
			sRet=sRet & "		<table>" & vbcrlf
			sRet=sRet & "			<tr style=""display:none"">" & vbcrlf
			sRet=sRet & "				<td><label for=""UserName"">�ǳƣ�</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""UserName"" type=""text"" id=""UserName"" size=""15"" maxlength=""0"" value="""" /></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
			sRet=sRet & "			<tr style=""display:none"">" & vbcrlf
			sRet=sRet & "				<td><label for=""Password"">���룺</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""Password"" type=""password"" id=""Password"" size=""15"" maxlength=""20"" value="""" />&nbsp;(�ο�������������)</td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
			sRet=sRet & "			<tr style=""display:none"">" & vbcrlf
			sRet=sRet & "				<td><label for=""Password"">��ҳ��</label></td>" & vbcrlf
			sRet=sRet & "				<td><input name=""homepage"" type=""text"" id=""homepage"" size=""42"" maxlength=""50"" value=""http://"" /></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
		End If
			sRet=sRet & "			<tr>" & vbcrlf
			sRet=sRet & "				<td><label for=""info"">���ݣ�</label></td>" & vbcrlf
			sRet=sRet & "				<td><textarea cols=""50"" rows=""6"" maxlength=""200"" name=""info"" id=""info""></textarea></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf
			sRet=sRet & "			<tr>" & vbcrlf
			sRet=sRet & "				<td></td>" & vbcrlf
			sRet=sRet & "				<td><input type=""submit"" id=""Submit"" value="" �� �� "" /></td>" & vbcrlf
			sRet=sRet & "			</tr>" & vbcrlf

		sRet=sRet & "		</table>" & vbcrlf
		sRet=sRet & "	</fieldset>" & vbcrlf
		sRet=sRet & "</form>" & vbcrlf
		JoinForm=sRet
	End Function
	'��׼ģ��
	Function AcceptJoin()
		Dim rs,sql,sRet,ustate
		'�ж��û��Ƿ��¼
		If oblog.checkuserlogined()=false Then
			AcceptJoin="<p>�������¼����ܽ����������</p>"
			Exit Function
		End If
		ustate=2
		'�жϼ�������
		Set rs=oblog.execute("select joinlimit,joinscores From oblog_team Where teamid="& CLng (g_id))
		If rs.Eof Then
			ErrMsg="Ŀ��" &oblog.CacheConfig(69)& "������!"
			Response.End
		End If
		select Case rs(0)
			Case -1
				ustate=3
			Case 1
				ErrMsg="��" &oblog.CacheConfig(69)& "ֻ�����鳤�������룬�����������"
			Case 2
				If oblog.l_uscores<rs(1) Then
					ErrMsg="���뱾" &oblog.CacheConfig(69)& "��Ҫ���� " & rs(1) & " �����,���Ļ��ֲ���"
				End If
			Case Else

		End select
		If ErrMsg<>"" Then
			AcceptJoin=ErrMsg
			Exit Function
		End If
		'�ж�֮ǰ�Ƿ��Ѽ��������
		Sql="select * From oblog_teamusers Where teamid=" & g_id & " And userid=" & oblog.l_uid
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open Sql,conn,1,3
		If Not rs.Eof Then
			select Case rs("state")
				Case 3
					sRet="���Ѿ��Ǹ�" &oblog.CacheConfig(69)& "�ĳ�Ա��"
				Case 1
					sRet="���Ѿ�������,��������ĺ�̨���ܻ�ܾ���" &oblog.CacheConfig(69)& "�����롣"
				Case 2
					sRet="���Ѿ���������,��ȴ�����Ա��ˡ�"
				Case 5
					sRet="���Ǹ�" &oblog.CacheConfig(69)& "����Ա,����Ҫ�������롣"
			End select
		Else
			'�ж�����
			rs.AddNew
			rs("teamid")=g_Id
			rs("userid")=oblog.l_uid
			rs("state")=ustate
			rs("info")= left(Request("info"),200)
			rs("icount")=0
			rs("addtime")=Now
			rs.Update
			if ustate=3 then
				sRet="���Ѽ����" &oblog.CacheConfig(69)& "��"
				'*&*&*
			else
				sRet="���ѳɹ���������,���ڵȴ�����Ա��ˡ�"
			end if
		End If
		rs.Close
		Set rs=Nothing
		AcceptJoin=sRet
	End Function

	'�޸���������ģ��
	Function LinksForm()
		Dim sRet
		sRet="<form action='group.asp?cmd=savelinks&gid="&g_Id&"' method='post' name='commentform' id='commentform' onSubmit='return Verifycomment()'>"& vbcrlf
		sRet=sRet & "<ul><p>���޸�������������</p></ul>"
		sRet=sRet & "<ul style=""display:none"">�ǳƣ�<input name='UserName' type='text' id='UserName' size='15' maxlength='20' value='' /></ul>" & vbcrlf
		sRet=sRet & "<ul style=""display:none;"">���룺<input name='Password' type='password' id='Password' size='15' maxlength='20' value='' /> (�ο�������������)</ul>" & vbcrlf
		sRet=sRet & "<ul style=""display:none;"">��ҳ��<input name='homepage' type='text' id='homepage' size='42' maxlength='50' value='http://' /></ul>"  & vbcrlf
		sRet=sRet & "<ul><input type='hidden' name='edit' id='edit' value='' />" & vbcrlf
		sRet=sRet & "<div id=""oblog_edit""></div> " & vbcrlf
		sRet=sRet & "</ul>" & vbcrlf
		sRet=sRet & "<ul><span id=""ob_code""></span><input type='submit' value=' �ύ '></ul>" & vbcrlf
		sRet=sRet & "</form></div>"& vbcrlf
		sRet=sRet & "<script src=""commentedit.asp""></script>"
		LinksForm=sRet
		sRet=""
	End Function
	Function SaveLinks()
		'�ж��Ƿ�Ϊ����Ա
		If IsManager=False Then
			ErrMsg= "��û��Ȩ�޽��д˲���"
			Response.End
		End If
		Dim rs,content
		content=Request.Form("oblog_edittext")
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "select * From oblog_team Where teamid=" & g_Id,conn,1,3
		rs("links")=oblog.Ubb_comment(EncodeJP(oblog.InterceptStr(oblog.filt_badword(content),250)))
		rs.Update
		rs.Close
		Set rs=Nothing
		Response.Redirect "group.asp?gid=" & g_id
	End Function
	'�޸�վ�㹫��ģ��
	Function PlacardForm()
		Dim sRet
		sRet="<div id=""form_comment""><form action='group.asp?cmd=saveplacard&gid="&g_Id&"' method='post' name='commentform' id='commentform' onSubmit='return Verifycomment()'>"& vbcrlf
		sRet=sRet & "<ul><p>���޸�����" &oblog.CacheConfig(69)& "����</p></ul>"
		sRet=sRet & "<ul style=""display:none"">�ǳƣ�<input name='UserName' type='text' id='UserName' size='15' maxlength='20' value='' /></ul>" & vbcrlf
		sRet=sRet & "<ul style=""display:none;"">���룺<input name='Password' type='password' id='Password' size='15' maxlength='20' value='' /> (�ο�������������)</ul>" & vbcrlf
		sRet=sRet & "<ul style=""display:none;"">��ҳ��<input name='homepage' type='text' id='homepage' size='42' maxlength='50' value='http://' /></ul>"  & vbcrlf
		sRet=sRet & "<ul><input type='hidden' name='edit' id='edit' value='' />" & vbcrlf
		sRet=sRet & "<div id=""oblog_edit""></div> " & vbcrlf
		sRet=sRet & "</ul>" & vbcrlf
		sRet=sRet & "<ul><span id=""ob_code""></span><input type='submit' value=' �ύ '></ul>" & vbcrlf
		sRet=sRet & "</form></div>"& vbcrlf
		sRet=sRet & "<script src=""commentedit.asp""></script>"
		PlacardForm=sRet
		sRet=""
	End Function
	Function SavePlacard()
		'�ж��Ƿ�Ϊ����Ա
		If IsManager=false Then
			ErrMsg= "��û��Ȩ�޽��д˲���"
			Response.End
		End If
		Dim rs,content
		content=Request.Form("oblog_edittext")
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "select * From oblog_team Where teamid=" & g_Id,conn,1,3
		rs("announce")=oblog.Ubb_comment(EncodeJP(oblog.InterceptStr(oblog.filt_badword(content),250)))
		rs.Update
		rs.Close
		Set rs=Nothing
		Response.Redirect "group.asp?gid=" & g_id
	End Function
	'��ɢģ��(�ݲ�����)

	function getminilist()
		Dim rs,Sql,sRet
		Sql="select top 10 topic,parentid,author,teamid,addtime From oblog_teampost Where iDepth=1 And teamid=" & g_id & " Order By postid Desc"
		Set rs=oblog.Execute(Sql)
		If rs.Eof Then

			sRet= sRet & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���»ظ� -->" & vbcrlf
			sRet= sRet & "<ul id=""GroupComments"">" & vbcrlf
			sRet= sRet & "	<li>��������</li>" & vbcrlf
			sRet= sRet & "</ul>" & vbcrlf
			sRet= sRet & "<!-- " &oblog.CacheConfig(69)& "���»ظ� end -->" & vbcrlf
		Else
			sRet= sRet & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "���»ظ� -->" & vbcrlf
			sRet= sRet & "<ul id=""GroupComments"">" & vbcrlf
			Do While Not rs.Eof
				sRet= sRet & "	<li><a href=""group.asp?gid="&rs(3)&"&pid="&rs(1)&""">" & OB_IIF(RemoveHtml(rs(0)),"����")&"</a><span class=""user"">"&rs(2)&"<span class=""time"">&nbsp;-&nbsp;"&rs(4)&"</span></span></li>" & vbcrlf
				rs.MoveNext
			Loop
			sRet= sRet & "</ul>" & vbcrlf
			sRet= sRet & "<!-- " &oblog.CacheConfig(69)& "���»ظ� end -->" & vbcrlf
		End if
		Set rs=Nothing
		getminilist=sRet
		sRet=""
	end function
	Function CheckQQLogin()
		Dim username,password
		username=oblog.filt_badstr(Trim(Request.form("username")))
		if username="" or oblog.strLength(username)>20 then oblog.adderrstr("���ֲ���Ϊ���Ҳ��ܴ���20���ַ���")
		if oblog.chk_badword(username)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
		password=Trim(Request.form("password"))
		if oblog.checkuserlogined()=false then
			password=md5(password)
			oblog.ob_chklogin username,password,0
		end if
	End Function
	'����б�
	Function ShowPhoto()
		Dim sRet,i,n
		Dim rs,lPage,lAll,lPages,sTitle,imgsrc
		Dim classid
		classid = Request("classid")
		If classid<>"" Then classid = CLng(classid) Else classid = 0
		G_P_Filename="group.asp?gid=" & g_id & "&cmd="&cmd&"&classid="&classid&"&page="
		sRet=""
		Set rs=Server.CreateObject("Adodb.RecordSet")
		if classid>0 then
			Sql = "select photo_path,fileID,photo_Title,a.userid,b.username,b.nickname from oblog_album a INNER JOIN oblog_user b ON a.userid=b.userid where TeamID="&g_id&" and sysClassId="&classid&"  order by photoID desc"
		else
			Sql = "select photo_path,fileID,photo_Title,a.userid,b.username,b.nickname from oblog_album a INNER JOIN oblog_user b ON a.userid=b.userid where TeamID="&g_id&"  order by photoID desc"
		end If
'		Response.Write(sql)
		rs.Open Sql,conn,1,1
'		Set rs=oblog.Execute(Sql)
'		sRet="<div id=""albumtop""><ul>"&GetSysClasses()&"<ul></div>"
		If rs.Eof Then
			rs.Close
			sRet= sRet & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "������Ƭ -->" & vbcrlf
			sRet= sRet & "<div>"
			sRet= sRet & "Ŀǰ��û���κ���Ƭ"
			sRet= sRet & "</div>" & vbcrlf
			sRet= sRet & "<!-- " &oblog.CacheConfig(69)& "������Ƭ end -->" & vbcrlf
			ShowPhoto=sRet
			sRet=""
			Exit Function
		End If
		'��ҳ
		If Request("page") = "" Or Request("page") ="0" then
			lPage = 1
		Else
			lPage = Int(Request("page"))
		End If
		lAll=Int(rs.recordcount)
		'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
		rs.CacheSize = iPage
		rs.PageSize = iPage
		rs.movefirst
		lPages = rs.PageCount
		If lPage>lPages Then lPage=lPages
		rs.AbsolutePage = lPage

		sRet= sRet & vbcrlf & "<!-- " &oblog.CacheConfig(69)& "������Ƭ -->"& vbcrlf
		sRet= sRet & "<div id=""GroupNewPhoto"">"& vbcrlf
		Do While Not rs.Eof And i < rs.PageSize
		'д����
				For n=1 to 4
					If Not rs.EOF Then
						If oblog.CacheConfig(67) = "1" Then
							imgsrc = "attachment.asp?path="&rs(0)
						Else
							imgsrc = ProIco(rs(0),3)
						End If
'						imgsrc=blogurl & rs(0)
						'imgsrc=Replace(imgsrc,right(imgsrc,3),"jpg")
						'imgsrc=Replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
						'if  not fso.FileExists(Server.MapPath(imgsrc)) then
							'imgsrc=blogurl&rsPhoto(0)
						'End if
						sRet=sRet&"	<ul class=""PhotoList"">" & vbcrlf
						sRet=sRet&"		<li class=""Photoimg""><a href=""group.asp?cmd=photocomment&gid=" & g_id & "&fileID="&rs(1)&""" title=""" & ob_IIF(rs(2), "�ޱ���") & """><img src='" & imgsrc & "'   /></a></li>" & vbcrlf
						sRet=sRet&"		<li class=""PhotoTitle""><a href=""group.asp?cmd=photocomment&gid=" & g_id & "&fileID="&rs(1)&""" title=""" & ob_IIF(rs(2),"�ޱ���") & """>" & ob_IIF(rs(2),"�ޱ���") & "</a></li>" & vbcrlf
						sRet=sRet&"		<li class=""Uploader""><a href=""go.asp?userid="&rs(3)&""" title="""&OB_IIF(rs("nickname"),rs("username"))&""">("&OB_IIF(rs("nickname"),rs("username"))&")</a></li>" & vbcrlf
						sRet=sRet&"	</ul>"& vbcrlf
						i=i+1
						rs.movenext
						if n>=iPage Then Exit For
					Else
					End if
				Next
		Loop
		sRet=sRet&"</div>"	& vbcrlf
		sRet=sRet&"<!-- " &oblog.CacheConfig(69)& "������Ƭ end -->" & vbcrlf
		rs.Close
		Set rs=Nothing
		'����һ���ײ���ҳ��
		If cmd<>"" Then sRet=sRet & "<div id=""GroupPages"">" & vbcrlf & PageBarNum(lAll,iPage,lPage,G_P_Filename) & vbcrlf & "</div>"
		'���ݽű�����
		'sRet= sRet & vbcrlf & "<div id=""comment_list""></div>"
		ShowPhoto=sRet
		sRet=""
	End Function

	'��ȡϵͳ����
	Function GetSysClasses()
		Dim rst,sReturn
		Set rst=oblog.Execute("select * From oblog_logclass Where idtype=1")
		If rst.Eof Then
			sReturn=""
		Else
			Do While Not rst.Eof
				sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
				rst.Movenext
			Loop
			sReturn = "<option value="""">��ѡ����Ƭ����</option><option value='0'>���з���</option>" & VBCRLF & sReturn
			sReturn="<select name=classid onchange=""javascript:window.location='group.asp?cmd=album&gid=" & g_id&"&classid='+this.options[this.selectedIndex].value;"">" & VBCRLF & sReturn & "</select>"
		End If
		rst.Close
		Set rst=Nothing
		sReturn=sReturn&"  <a href="""">Flash��ʽ���</a>"
		GetSysClasses = sReturn
	End Function
	'�������
	Function GetPhotoComment(ByRef sRet0)
		Dim trs,i
		Dim sPInfo
		Dim sRet,sRet1
		Dim rs,lPage,lAll,lPages,sTitle,imgsrc
		fileID = Request("fileid")
		If fileID <>"" Then fileID = CLng(fileID) Else fileID = 0
		Set trs = oblog.Execute ("select PHOTO_title,PHOTO_readme,PHOTO_path,fileID,photo_Name,addtime,b.username,b.nickname,b.userid FROM oblog_album a INNER JOIN oblog_user b ON a.userid=b.userid WHERE TeamID="&g_id&" AND fileid="&fileid)
		If TRS.EOF Then
			sRet0 = "����Ƭ������"
			trs.Close
			Set trs = Nothing
			Exit Function
		Else
			If oblog.CacheConfig(67) = "1" Then
				imgsrc = "attachment.asp?path="&trs("PHOTO_path")
			Else
				imgsrc = ProIco(trs(2),3)
			End If
			sPInfo = sPInfo & vbcrlf & "<div id=""GroupList"">" & vbcrlf
			sPInfo = sPInfo & "	<div class=""title""><a href="""&team_domain&""">"&g_Name&"</a> &gt; <a href=""group.asp?cmd=album&gid="&g_Id&""">������</a></div>" & vbcrlf
			sPInfo = sPInfo & "<div id=""Group_Photo_List"">" & vbcrlf
			sPInfo = sPInfo & "	<div class=""PhotoContent"">" & vbcrlf
			sPInfo = sPInfo & "		<div class=""PhotoTitle"">"&ob_IIF(trs(0),"�ޱ���")&"</div>" & vbcrlf
			sPInfo = sPInfo & "			<div class=""AddTime""><a href=""go.asp?userid="&trs("userid")&""" title=""����"&OB_IIF(trs("nickname"),trs("username"))&"�Ĳ���"" target=""_blank"">"&OB_IIF(trs("nickname"),trs("username"))&"</a> ������<span>"&trs("addtime")&"</span></div>" & vbcrlf
			sPInfo = sPInfo & "		<div class=""img"">" & vbcrlf
			sPInfo = sPInfo & "			<img src="""&imgsrc&""" onclick=""javascript:window.open(this.src);"" style=""cursor:pointer"" onload=""rsimg(this,500);"" alt=""����鿴ԭͼ""/>" & vbcrlf
			sPInfo = sPInfo & "		</div>" & vbcrlf
			sPInfo = sPInfo & "		<div class=""Content"">" & vbcrlf
			sPInfo = sPInfo & "			<div class=""ContentTitle"">ͼƬ��飺</div>" & vbcrlf
			sPInfo = sPInfo & ob_IIF(trs(1),"�޼��") & vbcrlf
			sPInfo = sPInfo & "		</div>" & vbcrlf
			sPInfo = sPInfo & "	</div>" & vbcrlf
		End If
		sTitle = ob_IIF(trs(0),"�ޱ���")
		g_show_title = sTitle
		sRet0 = sPInfo & "<div class=""Comments"">�������</div>" & vbcrlf
		G_P_Filename="group.asp?gid=" & g_id & "&cmd="&cmd&"&fileid="&fileid&"&page="
		sRet=""
		Set rs=Server.CreateObject("Adodb.RecordSet")
		SQL = "select a.*,b.user_icon1,b.username From oblog_albumcomment a,oblog_user b Where a.comment_user=b.username AND iState=1 AND MAINID="&fileid
		'OB_DEBUG (sql),1
		rs.Open SQL,conn,1,1
		If rs.Eof Then
			rs.Close
			sRet=sRet&"<div class=""Comments"">Ŀǰ��û���κ�����</div>" & vbcrlf
			sRet0=sPInfo&sRet
			sRet=""
		Else
			i=0
			'��ҳ
			If Request("page") = "" Or Request("page") ="0" then
				lPage = 1
			Else
				lPage = Int(Request("page"))
			End If
			lAll=Int(rs.recordcount)
			'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
			rs.CacheSize = iPage
			rs.PageSize = iPage
			rs.movefirst
			lPages = rs.PageCount
			If lPage>lPages Then lPage=lPages
			rs.AbsolutePage = lPage
			sRet1=""
			i=0
			Do While Not rs.Eof And i < rs.PageSize
				i=i+1
				sRet1 = sRet1 & vbcrlf &"	<div class=""CommentsContent"">" & vbcrlf
				sRet1 = sRet1 &"		<table>" & vbcrlf
				sRet1 = sRet1 &"			<tr>" & vbcrlf
				sRet1 = sRet1 &"				<td class=""t1"">" & vbcrlf
				sRet1 = sRet1 &"					<ul class=""User"">" & vbcrlf
				sRet1 = sRet1 &"						<li class=""userimg""><a href=""go.asp?user="&rs("COMMENT_USER")&""" title="""&rs("COMMENT_USER")&""" target=""_blank""><img src="""&OB_IIF(rs("user_icon1"),"images/ico_default.gif")&""" /></a></li>" & vbcrlf
				sRet1 = sRet1 &"						<li class=""username""><a href='go.asp?user="&rs("COMMENT_USER")&"' title="""&rs("COMMENT_USER")&""" target=""_blank"">"&rs("COMMENT_USER")&"</a></li>" & vbcrlf
				sRet1 = sRet1 &"					</ul>" & vbcrlf
				sRet1 = sRet1 &"				</td>" & vbcrlf
				sRet1 = sRet1 &"				<td  class=""t2"">" & vbcrlf
				sRet1 = sRet1 &"					<div class=""AddTime"">Posted <span id=""t_"&rs("COMMENTID")&""">"&rs("addtime")&"</span></div>" & vbcrlf
				sRet1 = sRet1 &"					<div class=""Content"">"&"<span id=""c_"&rs("COMMENTID")&""">"&oblog.Ubb_Comment(rs("COMMENT"))&"</span></div>" & vbcrlf
				sRet1 = sRet1 &"				</td>" & vbcrlf
				sRet1 = sRet1 &"			</tr>" & vbcrlf
				sRet1 = sRet1 &"		</table>" & vbcrlf
				sRet1 = sRet1 &"	</div>" & vbcrlf
				If imMode=1 or oblog.CodeCookie(rs("COMMENT_USER"))=Request.Cookies(cookies_name)("username") Then
					sRet1=Replace(sRet1,"$group_post_m$","<a href=""javascript:reply_quote('"& rs("COMMENTID")&"')"" >����</a><a href=""group.asp?cmd=del&gid=" & g_id &"&pid=" & rs("COMMENTID")& """>ɾ��</a>")
				Else
					sRet1=Replace(sRet1,"$group_post_m$","<a href=""javascript:reply_quote('"& rs("COMMENTID")&"')"" >����</a> ")
				End If
				rs.MoveNext
			Loop
		End If
		sRet1= sRet1 & "</div>"
		sRet1= sRet1 & "</div>"
		Dim sName
		sRet=sRet&CommentForm(fileID,1)
		sRet0 = sRet0 & sRet1&sRet
	End Function
End Class
%>