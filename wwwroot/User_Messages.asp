<!--#include file="user_top.asp"-->
<%
Dim action
action = Trim(Request("action"))
%>
<%
Dim sIp
Dim rs, sql, blog,sGuide
Dim id, usersearch, Keyword, sField
Keyword = Trim(Request("keyword"))
If Keyword <> "" Then
    Keyword = oblog.filt_badstr(Keyword)
End If
sIP=CheckIP(Request("IP"))
sField = Trim(Request("Field"))
usersearch = Trim(Request("usersearch"))
id = Trim(Request("id"))
If usersearch = "" Then
    usersearch = 0
Else
    usersearch = CLng(usersearch)
End If
G_P_FileName = "user_messages.asp?usersearch=" & usersearch

If Keyword <> "" Then
    G_P_FileName="user_messages.asp?usersearch=10&keyword="&Keyword&"&Field="&sField
End If
G_P_FileName =G_P_FileName & "&page="
If Request("page") <> "" Then G_P_This = Int(Request("page")) Else G_P_This = 1
If sIp <> "" Then G_P_FileName = "user_messages.asp"
select Case action
    Case "modify"
        Call modify
    Case "savemodify"
        Call savemodify
    Case "del"
        Call delmessage
	Case "passmessage"
		Call passmessage
    Case Else
        Call ClearIpMessages(sIp)
        Call main
End select
Set rs = Nothing
Set blog = Nothing
%>
	<tfoot>
		<tr>
			<td>
				<form name="form1" class="Search" id="MessagesSearch" action="user_messages.asp" method="get">
					<input type="hidden" name="t" value="<%=t%>">
					���ٲ������ԣ�&nbsp;
					<select size=1 name="usersearch" onChange="javascript:submit()">
						<option value="0">�г���������</option>
						<option value="2">����˵�����</option>
						<option value="3">����˵�����</option>
						<option value="4">�ҷ���������</option>
						<option value="10" selected>��ѡ���ѯ����</option>
					</select>
					&nbsp;&nbsp;������
					<select name="Field" id="Field">
						<option value="id">����</option>
						<option value="ip">����ip</option>
						<option value="topic" selected>���Ա���</option>
					</select>
					<input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
					<input type="submit" id="Submit" value="����">
				</form>
			</td>
		</tr>
	</tfoot>
</table>
</body>
</html>
<%
Sub main()
    Server.ScriptTimeOut = 999999999
    Dim  ssql,lPage,lAll,lPages,iPage,i,tsql
    ssql = "top 500 userid,messagetopic,addtime,messageid,message_user,addip,message,iState"
	tsql = " message_user = '"&oblog.l_uname&"' AND isguest = 0 "
    G_P_Guide = ""
    select Case usersearch
        Case 0
            sql="select "&ssql&" from [oblog_message] where userid="&oblog.l_uid&" AND isdel = 0 order by messageid desc"
            G_P_Guide = G_P_Guide & "����500������"
        Case 1
            sql="select "&ssql&" from [oblog_message] where userid="&oblog.l_uid&" AND isdel = 0 order by messageid desc"
            G_P_Guide = G_P_Guide & "�ҵ�����"
        Case 2
            sql="select "&ssql&" from [oblog_message] where userid="&oblog.l_uid&" and iState=1 AND isdel = 0 order by messageid desc"
            G_P_Guide = G_P_Guide & "����˵�����"
        Case 3
            sql="select "&ssql&" from [oblog_message] where userid="&oblog.l_uid&" and iState=0  AND isdel = 0 order by messageid desc"
            G_P_Guide = G_P_Guide & "����˵�����"
        Case 4
            sql="select "&ssql&" from [oblog_message] where "&tsql&" order by messageid desc"
            G_P_Guide = G_P_Guide & "�ҷ���������"
        Case 10
            If Keyword = "" Then
                oblog.adderrstr ("���󣺹ؼ��ֲ���Ϊ�գ�")
                oblog.showusererr
                Exit Sub
            Else
                select Case sField
                Case "id"
                    sql="select "&ssql&" from [oblog_message] where message_user like '%" & Keyword&"%' and userid="&oblog.l_uid&" order by messageid desc"
                    G_P_Guide = G_P_Guide & "���������л��к���<font color=red> " & Keyword & " </font>������"
                Case "topic"
                    sql="select "&ssql&" from [oblog_message] where messagetopic like '%" & Keyword & "%' and userid="&oblog.l_uid&" order by messageid desc"
                    G_P_Guide = G_P_Guide & "�����к��С� <font color=red>" & Keyword & "</font> ��������"
                Case "ip"
                    sql="select "&ssql&" from [oblog_message] where addip='" & Keyword&"' and userid="&oblog.l_uid&" order by messageid desc"
                    G_P_Guide = G_P_Guide & "����ipΪ<font color=red> " & Keyword & " </font>������"
                Case "content"
                    sql="select "&ssql&" from [oblog_message] where message like '%" & Keyword&"%' and userid="&oblog.l_uid&" order by messageid desc"
                    G_P_Guide = G_P_Guide & "���������а���<font color=red> " & Keyword & " </font>������"
                End select
            End If
        Case Else
    End select
    Set rs = Server.CreateObject("Adodb.RecordSet")
'	Response.Write sql
    rs.open sql, conn, 1, 3
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
					<li><a href="#" onclick="if (chk_idBatch(myform,'ͨ�����ѡ�е�������?')==true) {document.myform.action.value='passmessage';document.myform.iState.value='1'; document.myform.submit();}">ͨ�����</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ȡ�����ѡ�е�������?')==true) {document.myform.action.value='passmessage';document.myform.iState.value='0'; document.myform.submit();}">ȡ�����</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ɾ��ѡ�е�������?')==true) {document.myform.action.value='del';document.myform.submit();}">ɾ������</a></li>

					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="MessagesTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"><%=G_P_Guide%></td>
						<td class="t3">�ɣ�</td>
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
					<form name="myform" method="Post" action="user_messages.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
					<table id="Messages" class="TableList" cellpadding="0">
						<%
						'Do while not rs.EOF
						Do While Not rs.Eof And i < rs.PageSize
						i = i + 1
						%>
						<tr id="u<%=rs("messageid")%>" onclick="chk_iddiv('<%=rs("messageid")%>')">
							<td class="t1" title="���ѡ��">
								<input name='id' type='checkbox' id="c<%=rs("messageid")%>" value='<%=rs("messageid")%>' onclick="chk_iddiv('<%=rs("messageid")%>')" />
							</td>
							<td class="t2">
								<a href="go.asp?messageid=<%=rs("messageid")%>" target="_blank" title="cssbody=[dvbdy1] cssheader=[dvhdr1] body=[<%=oblog.filt_html(FilterUbb(rs("message")))%>]"><%If rs("iState") = 0 Then %>[����]<%End if%><%=oblog.filt_html(rs("messagetopic"))%></a><br />
								<span class="message_user">
									<%If rs("userid") = oblog.l_uid Then %>
										<%=oblog.filt_html(rs("message_user"))%>
									<%else%>
										<strong><%=oblog.filt_html(rs("message_user"))%></strong>
									<%end if%>
								</span>
								<!--ʱ��-->
								<div class="time">posted&nbsp;on&nbsp;<%=rs("addtime")%></div>
							</td>
							<td class="t3">
								<%=rs("addip")%>
							</td>
							<td class="t4">
								<%
								If rs("userid") = oblog.l_uid Then
									Response.write "<a href='user_messages.asp?action=modify&id=" & rs("messageid") & "&re=true'><span class=""blue"">�ظ�</span></a>&nbsp;"
									Response.write "<a href='user_messages.asp?action=modify&id=" & rs("messageid") & "'><span class=""green"">�޸�</span></a>&nbsp;"
									Response.write "<a href='user_messages.asp?action=del&id=" & rs("messageid") & "' onClick='return confirm(""ȷ��Ҫɾ����������"");'><span class=""red"">ɾ��</span></a>"
								End if
								%>
							</td>
						</tr>
						<%
							If i>iPage Then Exit Do
							rs.Movenext
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
        id = CLng (id)
    End If
    Set rsblog = Server.CreateObject("Adodb.RecordSet")
    sql="select * from [oblog_message] where messageid=" & id&" and userid="&oblog.l_uid
    rsblog.open sql, conn, 1, 1
    If rsblog.EOF Then
        rsblog.Close
        Set rsblog = Nothing
        oblog.adderrstr ("�����Ҳ���ָ�������ԣ�")
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
					<li><a href="#" onclick="purl('user_messages.asp','���Թ���')">���Թ���</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<form action="user_messages.asp?action=savemodify" method="post" name="oblogform" onSubmit="">
					<fieldset id="Messages" class="FieldsetForm">
						<%if Request("re") <> "true" then%>
						<legend>�޸�����</legend>
							<ul>
								<li><strong>���⣺</strong><input name="topic" type="text" value="<%=rsblog("messagetopic")%>" size="53" maxlength="30" /></li>
						<%else%>
						<legend>�ظ�����</legend>
							<ul>
								<li><strong>���⣺</strong><%=rsblog("messagetopic")%></li>
								<li><strong>���ߣ�</strong><%=rsblog("message_user")%></li>
								<li><div class="ubb_content"><table><tr><td><%=oblog.ubb_comment(rsblog("message"))%></td></tr></table></div></li>
						<%end if%>
								<li>
									<style type='text/css'>@import url('editor/ubb.css');</style>
									<Script src="editor/ubb.js"></Script>
									<div id="oblog_ubb">
										<div class="oblog_ubbtoolbar">
											<a href="javascript:InsertText(objActive,ReplaceText(objActive,'[B]','[/B]'),true);void(0)"><img src="images/bold.gif" alt="����"  border="0" align="absmiddle"></a>
											<a href="javascript:InsertText(objActive,ReplaceText(objActive,'[I]','[/I]'),true);void(0)"><img src="images/italic.gif" alt="б��" border="0" align="absmiddle" ></a>
											<a href="javascript:InsertText(objActive,ReplaceText(objActive,'[U]','[/U]'),true);void(0)"><img src="images/underline.gif" alt="�»���" border="0" align="absmiddle"></a>
											<a href="javascript:InsertText(objActive,ReplaceText(objActive,'[QUOTE]','[/QUOTE]'),true);void(0)"><img src="images/quote.gif" alt="��������" border="0" align="absmiddle"></a>
											<a href="javascript:UBB_smiley();void(0)"><img src="images/smiley.gif" alt="�������" border="0" align="absmiddle" id="A_smiley"></a>
										</div>
										<div id="oblog_ubbemot"></div>
										<textarea name="edit" cols="92" rows="8" id="oblog_edittext" class="oblog_ubbtext" ><%if rsblog("message")<>"" and Request("re") <> "true" then Response.Write Server.HtmlEncode(rsblog("message"))%></textarea>
									</div>
								</li>
								<li><input type="hidden" name="id" value="<%=rsblog("messageid")%>" /><input type="hidden" name="re" value="<%=request("re")%>" /><input type="submit" id="Submit" name="Submit2" value="ȷ���ύ" /></li>
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

Sub savemodify()
    Dim id, blog, userid
    id = CLng (Trim(Request("id")))
    sql="select * from oblog_message where messageid="&id&" and userid="&oblog.l_uid
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql, conn, 1, 3
    If Not rs.EOF Then
        userid = rs("userid")
        if Request("re")="true" then
			rs("message") = rs("message")&"[quote][b]����Ϊ"&oblog.filt_badword(oblog.l_uNickname)&"�Ļظ���[/b]"&vbcrlf&oblog.filt_badword(Request("edit"))&"[/quote]"
		else
			rs("message") = oblog.filt_badword(Request("edit"))
			rs("messagetopic") = oblog.InterceptStr(oblog.filt_badword(Trim(Request("topic"))), 250)
		end if
        rs.Update
        rs.Close
        Set blog = New class_blog
        blog.userid = userid
        blog.Update_message 0
        Set rs = Nothing
        Set blog = Nothing
    End If
    oblog.ShowMsg "�޸����Գɹ���", "user_messages.asp"
End Sub

Sub delmessage()
    If id = "" Then
        oblog.adderrstr ("������ָ��Ҫɾ�������ԣ�")
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        Dim n, i
        id = FilterIDs(id)
        n = Split(id, ",")
        For i = 0 To UBound(n)
            delonemessage (n(i))
        Next
    Else
        delonemessage (id)
    End If
    Set rs = Nothing
    oblog.ShowMsg "ɾ�����Գɹ�!", ""
End Sub

Sub delonemessage(id)
    id = CLng(id)
    Dim userid, messagefile,istate
    'Response.Write "idgfgfgfgf" & id
  	'Response.End
    sql = "select * from [oblog_message] where messageid=" & CLng (id) & " and userid=" & oblog.l_uId
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql, conn, 1, 3
    If Not rs.EOF Then
        userid = rs("userid")
        messagefile = rs("messagefile")
		istate=rs("istate")
        rs.Delete
        rs.Close
        Set blog = New class_blog
        blog.userid = userid
        'blog.update_message 0,0,0,""
        blog.Update_message 0
        blog.Update_newmessage userid
        If istate = 1 Then
        	oblog.execute("update [oblog_user] set message_count=message_count-1,scores=scores-" & oblog.CacheScores(5)&" where userid="&userid)
        End If
        Set blog = Nothing
    Else
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("������ɾ��Ȩ�ޣ�")
        oblog.showusererr
        Exit Sub
    End If
End Sub

Sub ClearIpMessages(sIp)
    If sIp <> "" Then oblog.Execute ("Delete From oblog_message Where addIp='" & sIp & "' and userid=" & oblog.l_uId)
End Sub

Function FilterUbb(byval strHTML)
	Dim objRegExp, strOutput
	Set objRegExp = New Regexp
	strOutput=strHTML
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern="(\[EMOT\])(.[^\[]*)(\[\/EMOT\])"
	strOutput = objRegExp.replace(strOutput, "")
	objRegExp.Pattern =  "\[[^\]]*\]"
	strOutput = objRegExp.replace(strOutput, " ")
	FilterUbb = strOutput
	Set objRegExp = Nothing
End Function


Sub passmessage()
	Dim iState
	iState=Request("iState")
    If id = "" Then
        oblog.adderrstr ("������ָ��Ҫ��˵����ԣ�")
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        Dim n, i
        id = FilterIDs(id)
        n = Split(id, ",")
        For i = 0 To UBound(n)
            passonemessage n(i),iState
        Next
    Else
        passonemessage id,iState
    End If
    Set rs = Nothing
    oblog.ShowMsg "������Գɹ�!", ""
End Sub

Sub passonemessage(id,iState)
    id = CLng(id)
	iState= CLng(iState)
    Dim userid
	Dim sScore,tstr
	If iState = 1 Then
		sScore=oblog.CacheScores(5)
		tstr = "+"
	Else
		sScore=-1*Abs(oblog.CacheScores(5))
		tstr = "-"
	End if
    sql = "select * from [oblog_message] where messageid=" & CLng(id) & " and userid=" & oblog.l_uId
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql, conn, 1, 3
    If Not rs.EOF Then
        userid = rs("userid")
		If rs("iState")<>iState Then
'			If oblog.CacheConfig(50) = 0 Then
				oblog.GiveScore "",sScore,""
				oblog.execute("update oblog_user set message_count=message_count"&tstr&"1 where userid="&userid)
				oblog.execute("update oblog_setup set message_count=message_count"&tstr&"1")
'			End If
		Else
			Exit Sub
		End if
        rs("iState")=iState
        rs.Update
        rs.Close
        Set blog = New class_blog
        blog.userid = userid
        'blog.update_message 0,0,0,""
        blog.Update_message 0
        blog.Update_newmessage userid
        Set blog = Nothing
    Else
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("�����޲���Ȩ�ޣ�")
        oblog.showusererr
        Exit Sub
    End If
End Sub
%>