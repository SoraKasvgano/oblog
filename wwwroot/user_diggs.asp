<!--#include file="user_top.asp"-->
<%
Dim action
action = Trim(Request("action"))
%>
<%
Dim sIp, sGuide
Dim rs, sql,mainid
Dim id, cmd, Keyword, sField

Keyword = Trim(Request("keyword"))
If Keyword <> "" Then
    Keyword = oblog.filt_badstr(Keyword)
End If
sField = Trim(Request("Field"))
cmd = Trim(Request("cmd"))

id = Trim(Request("id"))
mainid = CLng(Request("mainid"))
sIp = CheckIP(Request("IP"))
If cmd = "" Then
    cmd = 0
Else
    cmd = Int(cmd)
End If
G_P_FileName = "user_diggs.asp?cmd=" & cmd
If Keyword <> "" Then
    G_P_FileName="user_diggs.asp?cmd=10&keyword="&Keyword&"&Field="&sField
End If
G_P_FileName =G_P_FileName & "&page="
If Request("page") <> "" Then G_P_This = Int(Request("page")) Else G_P_This = 1
If sIp <> "" Then G_P_FileName = "user_diggs.asp"

If action = "modify" Then
    Call modify
ElseIf action = "savemodify" Then
    Call Savemodify
ElseIf action = "del" Then
    Call delcomment
ElseIf action = "passcomment" Then
	Call passcomment
Else
    Call main
End If
%>
<%
Sub main()
    Server.scriptTimeOut = 999999999
    Dim  ssql,i,lPage,lAll,lPages,iPage
    ssql = "top 500 authorid,diggtitle,diggurl,addtime,diggID,classid,diggdes,author,iState,diggnum"
    sGuide = ""
    select Case cmd
        Case 0
            sql="select "&ssql&" from [oblog_userdigg] where authorid="&oblog.l_uid&"  order by diggID desc"
            sGuide = sGuide & "����500ƪ�û��Ƽ���־"
        Case 2
            sql="select "&ssql&" from [oblog_userdigg] where authorid="&oblog.l_uid&" and iState=1 order by diggID desc"
            sGuide = sGuide & "����˵��û��Ƽ���־"
        Case 3
            sql="select "&ssql&" from [oblog_userdigg] where authorid="&oblog.l_uid&" and iState=0 order by diggID desc"
            sGuide = sGuide & "����˵��û��Ƽ���־"
		Case 4
            sql="select top 500 a.authorid,diggtitle,diggurl,a.addtime,a.diggID,classid,diggdes,author,iState,diggnum,userid from [oblog_userdigg] a ,oblog_digg b where userid="&oblog.l_uid&" AND a.diggid=b.diggid  AND iState=1  order by a.diggID desc"
            sGuide = sGuide & "���Ƽ�����־"
        Case 10
            If Keyword = "" Then
                oblog.adderrstr ("���󣺹ؼ��ֲ���Ϊ�գ�")
                oblog.showusererr
                Exit Sub
            Else
                select Case sField
                Case "topic"
                    sql="select "&ssql&" from [oblog_userdigg] where diggtitle like '%" & Keyword & "%' and authorid="&oblog.l_uid&" order by diggID desc"
                    sGuide = sGuide & "�����к��С� <font color=red>" & Keyword & "</font> �����û��Ƽ���־"
                End select
            End If
        Case Else
    End select
    Set rs = Server.CreateObject("Adodb.RecordSet")
'	OB_DEBUG sql,1
    rs.Open sql, conn, 1, 3
    lAll=Int(rs.recordcount)
    If lAll=0 Then
    	rs.Close
    	Set rs=Nothing
    	%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<!-- û����ؼ�¼ -->
					<div class="msg"><%=sGuide & " û����ؼ�¼" %></div>
					<!-- û����ؼ�¼ end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/72.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
    	<%
    	Exit Sub
    End If
    i=0
    iPage=20
	'��ҳ
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = Int(Request("page"))
	End If

	'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
	rs.CacheSize = iPage
	rs.PageSize = iPage
	rs.movefirst
	lPages = rs.PageCount
	If lPage>lPages Then lPage=lPages
	rs.AbsolutePage = lPage
	i=0
	%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="chk_idAll(myform,1);">ȫ��ѡ��</a></li>
					<li><a href="#" onclick="chk_idAll(myform,0);">ȫ��ȡ��</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ͨ�����ѡ�еļ�¼��?')==true) {document.myform.action.value='passcomment';document.myform.iState.value='1'; document.myform.submit();}">ͨ�����</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ȡ�����ѡ�еļ�¼��?')==true) {document.myform.action.value='passcomment';document.myform.iState.value='0'; document.myform.submit();}">ȡ�����</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ɾ��ѡ�еļ�¼��?')==true) {document.myform.action.value='del'; document.myform.submit();}">ɾ��</a></li>

					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="Diggstop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"><%=sGuide%></td>
						<td class="t3">�Ƽ���</td>
						<td class="t4">����</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<form name="myform" method="Post" action="user_diggs.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
					<table id="Diggs" class="TableList" cellpadding="0">
						<%
						'Do while not rs.EOF
						Do While Not rs.Eof And i < rs.PageSize
						i = i + 1
						%>
						<tr id="u<%=rs("diggID")%>"  onclick="chk_iddiv('<%=rs("diggID")%>')">
							<td class="t1" title="���ѡ��">
								<input name='id' type='checkbox' id="c<%=rs("diggID")%>" value='<%=rs("diggID")%>' onclick="chk_iddiv('<%=rs("diggID")%>')" />
							</td>
							<td class="t2">
								<a href="<%=rs("diggurl")%>" target="_blank" title="cssbody=[dvbdy1] cssheader=[dvhdr1] body=[<%=oblog.filt_html(rs("diggdes"))%>]"><%If rs("iState") = 0 Then %>[����]<%End if%><%=oblog.filt_html(rs("diggtitle"))%></a><br />
								<span class="message_user">posted by
									<%If rs("authorid") = oblog.l_uid Then %>
										<%=rs("author")%>
									<%else%>
										<strong><%=rs("author")%></strong>
									<%End if%>
								</span>
								<!--ʱ��-->
								<div class="time">&nbsp;&nbsp;<%=rs("addtime")%></div>
							</td>
							<td class="t3">
								<%=OB_IIF(rs("diggnum"),0)%>
							</td>
							<td class="t4">
								<%If rs("authorid") = oblog.l_uid Then %>
									<a href="user_diggs.asp?action=modify&id=<%=rs("diggID")%>"><span class="green">�޸�</span></a>&nbsp;
									<a href="user_diggs.asp?action=del&id=<%=rs("diggID")%>" onClick="return confirm('ȷ��Ҫɾ�����Ƽ�������');"><span class="red">ɾ��</span></a>
								<%Else%>
									<span class="red">��Ȩ����</span>
								<%End if%>
							</td>
						</tr>
						<%
							If i>iPage Then Exit Do
							rs.movenext
							Loop
							rs.Close
							Set rs = Nothing
						%>
					</table>
					<input type="hidden" name="iState">
					<input type="hidden" name="action" value="">
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/90.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%
End Sub
%>
	<tfoot>
		<tr>
			<td>
				<form class="Search" id="CommentsSearch" action="user_diggs.asp" name="form1" method="get">
					<input type="hidden" name="t" value="<%=t%>">
					���ٲ����Ƽ����£�&nbsp;
					<select size=1 name="cmd" onChange="javascript:submit()">
						<option value="0">�г������Ƽ�����</option>
						<option value="2">����˵��Ƽ�����</option>
						<option value="3">����˵��Ƽ�����</option>
						<option value="4">���Ƽ�������</option>
						<option value="10" selected>��ѡ���ѯ����</option>
					</select>
					&nbsp;&nbsp;������
					<select name="Field" id="Field">
						<option value="topic" selected>����</option>
					</select>
					<input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
					<input type="submit" id="Submit" value="����" />
				</form>
			</td>
		</tr>
	</tfoot>
</table>
</body>
</html>
<%
Sub modify()
    Dim id
    Dim rsblog, sql
    Dim restr
    id = Trim(Request("id"))
    If id = "" Then
        oblog.adderrstr ("���󣺲������㣡")
        oblog.showusererr
        Exit Sub
    Else
        id = CLng(id)
    End If
	id=FilterIds(id)
    Set rsblog = Server.CreateObject("Adodb.RecordSet")
    sql="select * from [oblog_userdigg] where diggID=" & id&" and authorid="&oblog.l_uid
    rsblog.Open sql, conn, 1, 1
    If rsblog.EOF Then
        rsblog.Close
        Set rsblog = Nothing
        oblog.adderrstr ("������Ȩ�ޣ�ֻ��blog���˲��ܲ�����")
        oblog.showusererr
        Exit Sub
    End If
%>
<SCRIPT language=javascript>
var ubbimg='';
</SCRIPT>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="purl('user_diggs.asp','DIGG')">�Ƽ����¹���</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<form action="user_diggs.asp?action=savemodify" method="post" name="oblogform" onSubmit="">
					<fieldset id="Comments" class="FieldsetForm">
						<legend>�޸�</legend>
							<ul>
								<li><strong>���⣺</strong><input name="topic" type="text" value="<%=rsblog("diggtitle")%>" size="53" maxlength="30" /></li>

								<li><strong>ժҪ��</strong>(��֧��HTML)<br />
										<textarea name="edit" cols="92" rows="6" id="oblog_edittext" class="oblog_ubbtext" ><%if rsblog("diggdes")<>"" then response.Write Server.HtmlEncode(rsblog("diggdes"))%></textarea>
								</li>
								<li><input type="hidden" name="id" value="<%=rsblog("diggID")%>" /><input type="submit" id="Submit" name="Submit2" value="ȷ���ύ" /></li>
							</ul>
					</fieldset>
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/72.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%
    rsblog.Close
    Set rsblog = Nothing
End Sub

Sub Savemodify()
    Dim id, rsblogchk, blog, logid, uid,des,topic
    id = CLng(Trim(Request("id")))
	des = oblog.InterceptStr(oblog.filt_badword(RemoveHtml(Trim(Request("edit")))),255)
	topic = oblog.InterceptStr(oblog.filt_badword(Trim(Request("topic"))), 255)
    sql="select * from oblog_userdigg where diggID="&id&" and authorid="&oblog.l_uid
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 1, 3
	If Not rs.EOF Then
		uid = rs("authorid")
		logid = rs("logid")
		rs("diggdes") = des
		rs("diggtitle") = topic
		rs.Update
		rs.Close
		Set rs = Nothing
		Set rs = Server.CreateObject("adodb.recordset")
		  sql = "select * from oblog_log where logid="&logid&""
		  rs.Open sql ,conn,1,3
		  rs("Abstract") = des
		  rs("topic") = topic
		  rs.Update
		  rs.close
		  Set rs = Nothing
		Set blog = New class_blog
		blog.userid = uid
		blog.Update_log logid, 0
		blog.update_index 0
		Set blog = Nothing
		oblog.ShowMsg "�޸ĳɹ���", "user_diggs.asp"
	Else
        oblog.adderrstr ("�������޸�Ȩ�ޣ�")
        oblog.showusererr
	End if
End Sub


Sub delcomment()
    Dim blog
    If id = "" Then
        oblog.adderrstr ("������ָ��Ҫɾ�����Ƽ����£�")
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        id = FilterIDs(id)
        Dim n, i
        n = Split(id, ",")
        For i = 0 To UBound(n)
            delonecomment (n(i))
        Next
    Else
        delonecomment (id)
    End If
    oblog.ShowMsg "ɾ���Ƽ����³ɹ�!", ""
End Sub

Sub delonecomment(id)
	On Error Resume Next
    Dim blog
    id = CLng(id)
	id=FilterIds(id)
    Dim uid, mainid,trs
    sql = "select authorid,logid from [oblog_userdigg] where diggID=" & CLng(id) & " and authorid=" & oblog.l_uId
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 1, 3
	'��ɾ�����ݿ��д�����¼
    If Not rs.EOF Then
        uid = rs(0)
        mainid = rs(1)
 		rs.Delete
        rs.Close
    Else
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("������ɾ��Ȩ�ޣ�")
        oblog.showusererr
        Exit Sub
    End If
	'����������¼���������־��¼ȡ����,����DIGGNUM����Ϊ0
	Set trs =  Server.CreateObject("adodb.recordset")
	trs.open "SELECT b.diggnum FROM oblog_digg a INNER JOIN oblog_log b ON a.logid = b.logid WHERE a.diggID = " & CLng(id),CONN,1,3
	If Not trs.Eof Then
		While Not trs.EOF
			trs(0) = 0
			trs.Update
			trs.MoveNext
		Wend
	End If
'	oblog.Execute ("UPDATE b SET diggnum = 0  FROM oblog_digg AS a INNER JOIN oblog_log AS b ON a.logid = b.logid WHERE a.diggID =" & Int(id))
	'�����û����ӵĻ��ֿ�ȥ
	Set trs = oblog.Execute ("SELECT COUNT(DID) FROM oblog_digg WHERE diggID = " & CLng(id))
	If Not trs.Eof Then oblog.GiveScore "",-1*Abs(oblog.CacheScores(22))*trs(0),""
	oblog.Execute ("UPDATE oblog_user SET diggs = diggs - "&trs(0)&"  WHERE userid = " & oblog.l_uId)
	'�����û���DIGGNUM��ȥ����ó�����ֵ
	oblog.Execute ("DELETE FROM oblog_digg WHERE diggID = " & CLng (id))
	'���йؼ�¼ȫ��ɾ��
	Set trs = Nothing
End Sub

Sub passcomment()
	Dim iState
	iState=Request("iState")
    Dim blog
    If id = "" Then
        oblog.adderrstr ("������ָ��Ҫ��˵��Ƽ����£�")
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        id = FilterIDs(id)
        Dim n, i
        n = Split(id, ",")
        For i = 0 To UBound(n)
            passonecomment n(i),iState
        Next
    Else
        passonecomment id,iState
    End If
    oblog.ShowMsg "����Ƽ����³ɹ�!", ""
End Sub

Sub passonecomment(id,iState)
    Dim blog
	iState=CLng(iState)
    sql = "select iState from [oblog_userdigg] where diggID=" & CLng (id) & " and authorid=" & oblog.l_uId
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then
		If rs("iState")=iState Then	Exit Sub
		rs("iState")=iState
        rs.Update
        rs.Close
    Else
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("�����޲���Ȩ�ޣ�")
        oblog.showusererr
        Exit Sub
    End If
End Sub
%>