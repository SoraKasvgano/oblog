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
G_P_FileName = "User_Albumcomments.asp?cmd=" & cmd
If Keyword <> "" Then
    G_P_FileName="User_Albumcomments.asp?cmd=10&keyword="&Keyword&"&Field="&sField
End If
G_P_FileName =G_P_FileName & "&page="
If Request("page") <> "" Then G_P_This = Int(Request("page")) Else G_P_This = 1
If sIp <> "" Then G_P_FileName = "User_Albumcomments.asp"

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
<script language=javascript>

function VerifySubmit()
{
    topic = del_space(document.oblogform.topic.value);
     if (topic.length == 0)
     {
        alert("��������д��Ŀ!");
    return false;
     }

    submits();
    if (document.oblogform.edit.value == "")
     {
        alert("����������!");
    return false;
     }
    return true;
}
</script>
<%
Sub main()
    Server.scriptTimeOut = 999999999
    Dim  ssql,i,lPage,lAll,lPages,iPage,tsql
    ssql = "top 500 userid,mainid,commenttopic,addtime,commentid,comment_user,addip,comment,iState"
	tsql = " comment_user = '"&oblog.l_uname&"' AND isguest = 0 AND isdel = 0  or isdel is null"
    sGuide = ""
    select Case cmd
        Case 0
            sql="select "&ssql&" from [oblog_Albumcomment] where userid="&oblog.l_uid&" AND isdel = 0  or isdel is null order by commentid desc"
            sGuide = sGuide & "����500ƪ����"
        Case 1
            sql="select "&ssql&" from [oblog_Albumcomment] where userid="&oblog.l_uid&" AND isdel = 0  or isdel is null order by commentid desc"
            sGuide = sGuide & "�������������"
        Case 2
            sql="select "&ssql&" from [oblog_Albumcomment] where userid="&oblog.l_uid&" and iState=1 AND isdel = 0 or isdel is null order by commentid desc"
            sGuide = sGuide & "����˵�����"
        Case 3
            sql="select "&ssql&" from [oblog_Albumcomment] where userid="&oblog.l_uid&" and iState=0  AND isdel = 0 or isdel is null order by commentid desc"
            sGuide = sGuide & "����˵�����"
        Case 4
            sql="select "&ssql&" from [oblog_Albumcomment] where "&tsql&" order by commentid desc"
            sGuide = sGuide & "�ҷ���������"
        Case 10
            If Keyword = "" Then
                oblog.adderrstr ("���󣺹ؼ��ֲ���Ϊ�գ�")
                oblog.showusererr
                Exit Sub
            Else
                select Case sField
                Case "id"
                    sql="select "&ssql&" from [oblog_Albumcomment] where comment_user like '%" & Keyword&"%' and userid="&oblog.l_uid&" order by commentid desc"
                    sGuide = sGuide & "�������ƺ���<font color=red> " & Keyword & " </font>������"
                Case "topic"
                    sql="select "&ssql&" from [oblog_Albumcomment] where commenttopic like '%" & Keyword & "%' and userid="&oblog.l_uid&" order by commentid desc"
                    sGuide = sGuide & "�����к��С� <font color=red>" & Keyword & "</font> ��������"
                Case "ip"
                    sql="select "&ssql&" from [oblog_Albumcomment] where addip='" & Keyword&"' and userid="&oblog.l_uid&" order by commentid desc"
                    sGuide = sGuide & "����ip����<font color=red> " & Keyword & " </font>������"
                Case "content"
                    sql="select "&ssql&" from [oblog_Albumcomment] where comment like '%" & Keyword&"%'  and userid="&oblog.l_uid&" order by commentid desc"
                    sGuide = sGuide & "���������а���<font color=red> " & Keyword & " </font>������"
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
					<li><a href="#" onclick="if (chk_idBatch(myform,'ͨ�����ѡ�е�������?')==true) {document.myform.action.value='passcomment';document.myform.iState.value='1'; document.myform.submit();}">ͨ�����</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ȡ�����ѡ�е�������?')==true) {document.myform.action.value='passcomment';document.myform.iState.value='0'; document.myform.submit();}">ȡ�����</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ɾ��ѡ�е�������?')==true) {document.myform.action.value='del'; document.myform.submit();}">ɾ������</a></li>

					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="AlbumCommentsTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"><%=sGuide%></td>
						<td class="t3">�����ߣɣ�</td>
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
					<form name="myform" method="Post" action="User_Albumcomments.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
					<table id="AlbumComments" class="TableList" cellpadding="0">
						<%
						'Do while not rs.EOF
						Do While Not rs.Eof And i < rs.PageSize
						i = i + 1
						%>
						<tr id="u<%=rs("commentid")%>"  onclick="chk_iddiv('<%=rs("commentid")%>')">
							<td class="t1" title="���ѡ��">
								<input name='id' type='checkbox' id="c<%=rs("commentid")%>" value='<%=rs("commentid")%>' onclick="chk_iddiv('<%=rs("commentid")%>')" />
							</td>
							<td class="t2">
								<a href="go.asp?fileid=<%=rs("mainid")%>#<%=rs("commentid")%>" target="_blank" title="cssbody=[dvbdy1] cssheader=[dvhdr1] body=[<%=oblog.filt_html(FilterUbb(rs("comment")))%>]"><%If rs("iState") = 0 Then %>[����]<%End if%><%=oblog.filt_html(rs("commenttopic"))%></a><br />
								<span class="message_user">
									<%If rs("userid") = oblog.l_uid Then %>
										<%=rs("comment_user")%>
									<%else%>
										<strong><%=rs("comment_user")%></strong>
									<%End if%>
								</span>
								<!--ʱ��-->
								<div class="time">post&nbsp;by&nbsp;<%=rs("addtime")%></div>
							</td>
							<td class="t3">
								<%=rs("addip")%>
							</td>
							<td class="t4">
								<%If rs("userid") = oblog.l_uid Then %>
									<a href="User_Albumcomments.asp?action=modify&id=<%=rs("commentid")%>&re=true"><span class="blue">�ظ�</span></a>&nbsp;
									<a href="User_Albumcomments.asp?action=modify&id=<%=rs("commentid")%>"><span class="green">�޸�</span></a>&nbsp;
									<a href="User_Albumcomments.asp?action=del&id=<%=rs("commentid")%>" onClick="return confirm('ȷ��Ҫɾ����������');"><span class="red">ɾ��</span></a>
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
				<form class="Search" id="CommentsSearch" action="User_Albumcomments.asp" name="form1" method="get">
					<input type="hidden" name="t" value="<%=t%>">
					���ٲ������ۣ�&nbsp;
					<select size=1 name="cmd" onChange="javascript:submit()">
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
						<option value="topic" selected>���۱���</option>
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
    Set rsblog = Server.CreateObject("Adodb.RecordSet")
    sql="select * from [oblog_Albumcomment] where commentid=" & id&" and userid="&oblog.l_uid
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
					<li><a href="#" onclick="purl('User_Albumcomments.asp','���۹���')">���۹���</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<form action="User_Albumcomments.asp?action=savemodify" method="post" name="oblogform" onSubmit="">
					<fieldset id="Comments" class="FieldsetForm">
						<%if Request("re") <> "true" then%>
						<legend>�޸�����</legend>
							<ul>
								<li><strong>���⣺</strong><input name="topic" type="text" value="<%=rsblog("commenttopic")%>" size="53" maxlength="30" /></li>
						<%else%>
						<legend>�ظ�����</legend>
							<ul>
								<li><strong>���⣺</strong><%=rsblog("commenttopic")%></li>
								<li><strong>���ߣ�</strong><%=rsblog("comment_user")%></li>
								<li><div class="ubb_content"><table><tr><td><%=oblog.ubb_comment(rsblog("comment"))%></td></tr></table></div></li>
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
										<textarea name="edit" cols="92" rows="6" id="oblog_edittext" class="oblog_ubbtext" ><%if rsblog("comment")<>"" and Request("re") <> "true" then response.Write Server.HtmlEncode(rsblog("comment"))%></textarea>
									</div>
								</li>
								<li><input type="hidden" name="id" value="<%=rsblog("commentid")%>" /><input type="hidden" name="re" value="<%=request("re")%>" /><input type="submit" id="Submit" name="Submit2" value="ȷ���ύ" /></li>
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
    Dim id, rsblogchk, blog, logid, uid
    id = CLng(Trim(Request("id")))
    sql="select * from oblog_Albumcomment where commentid="&id&" and userid="&oblog.l_uid
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 1, 3
    uid = rs("userid")
    logid = rs("mainid")
   if Request("re")="true" then
		rs("comment") = rs("comment")&"[quote][b]����Ϊ"&oblog.filt_badword(oblog.l_uNickname)&"�Ļظ���[/b]"&vbcrlf&oblog.filt_badword(Request("edit"))&"[/quote]"
	else
		rs("comment") = oblog.filt_badword(Request("edit"))
		rs("commenttopic") = oblog.InterceptStr(oblog.filt_badword(Trim(Request("topic"))), 250)
	end if
    rs.Update
    rs.Close
    Set rs = Nothing
    Set blog = New class_blog
    blog.userid = uid
    blog.Update_log logid, 0
    Set blog = Nothing
    oblog.ShowMsg "�޸����۳ɹ���", "User_Albumcomments.asp"
End Sub


Sub delcomment()
    Dim blog, rstComment
    If id = "" Then
        oblog.adderrstr ("������ָ��Ҫɾ�������ۣ�")
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
    oblog.ShowMsg "ɾ�����۳ɹ�!", ""
End Sub

Sub delonecomment(id)
    Dim  rstComment, CommentNum
    id = CLng(id)
    Dim uid, mainid,istate
    Set rstComment=Server.CreateObject("Adodb.Recordset")
    sql = "select userid,mainid,istate from [oblog_Albumcomment] where commentid=" & CLng(id) & " and userid=" & oblog.l_uId
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then
        uid = rs(0)
        mainid = rs(1)
		istate=rs(2)
 		rs.Delete
        rs.Close
        '���¼���������Ŀ
        Set rstComment = Server.CreateObject("adodb.recordset")
        rstComment.Open "select Count(commentid) From [oblog_Albumcomment] Where mainid=" & CLng (mainid), conn, 1, 1
        If rstComment.EOF Then
            CommentNum = 0
        Else
            If IsNull(rstComment(0)) Or Not IsNumeric(rstComment(0)) Then
                CommentNum = 0
            Else
                CommentNum = rstComment(0)
            End If
        End If
        rstComment.Close
        'oblog.Execute ("update [oblog_log] set commentnum=" & CommentNum & ",scores=scores-" & oblog.CacheScores(6) & " where logid=" & mainid)
        rstComment.Open "select commentnum From [oblog_album] where fileid=" & mainid,conn,1,3
        rstComment(0)=CommentNum
        rstComment.Update
        rstComment.Close
        'oblog.Execute ("update [oblog_user] set comment_count=comment_count-1,scores=scores-" &  oblog.CacheScores(6) & " where userid=" & uid)
        rstComment.Open "select comment_count,scores From [oblog_user] where userid=" & uid,conn,1,3
        rstComment(0)=rstComment(0)-1
        If istate=1 Then
	        If rstComment(1)>Int(oblog.CacheScores(6)) Then
	        	rstComment(1)=rstComment(1)-Int(oblog.CacheScores(6))
	        Else
	        	rstComment(1)=0
	        End If
	    End If
        rstComment.Update
        Set rstComment = Nothing
    Else
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("������ɾ��Ȩ�ޣ�")
        oblog.showusererr
        Exit Sub
    End If
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

Sub passcomment()
	Dim iState
	iState=Request("iState")
    Dim blog, rstComment
    If id = "" Then
        oblog.adderrstr ("������ָ��Ҫ��˵����ۣ�")
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
    oblog.ShowMsg "������۳ɹ�!", ""
End Sub

Sub passonecomment(id,iState)
    Dim blog
    id = CLng(id)
	iState=CLng(iState)
    Dim uid, mainid
	Dim sScore,tstr
	If iState = 1 Then
		sScore=oblog.CacheScores(6)
		tstr = "+"
	Else
		sScore=-1*Abs(oblog.CacheScores(6))
		tstr = "-"
	End if
    sql = "select userid,mainid,iState from [oblog_Albumcomment] where commentid=" & CLng (id) & " and userid=" & oblog.l_uId
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then
        uid = rs(0)
        mainid = rs(1)
		If rs("iState")<>iState Then
'			If oblog.CacheConfig(50) = 0 Then
				oblog.GiveScore "",sScore,""
				oblog.execute("update oblog_user set comment_count=comment_count"&tstr&"1 where userid="&uid)
				oblog.execute("update oblog_setup set comment_count=comment_count"&tstr&"1")
'			End If
		Else
			Exit Sub
		End if
		rs("iState")=iState
        rs.Update
        rs.Close
        Set blog = New class_blog
        blog.userid = uid
		blog.Update_log mainid, 0
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