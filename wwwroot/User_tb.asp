<!--#include file="user_top.asp"-->
<%
Dim action
action = Trim(Request("action"))
%>
</head>
<body scroll="no" style="overflow:auto!important;overflow:hidden;background:#fff;">
<%
Dim  sGuide
Dim rs, sql,mainid
Dim id, cmd
cmd = Trim(Request("cmd"))
id = Request("id")
If cmd = "" Then
    cmd = 0
Else
    cmd = Int(cmd)
End If
G_P_FileName = "user_tb.asp?cmd=" & cmd & "&page="

If  action = "del" Then
    Call deltb
Else
    Call main
End If
%>
</table>
</body>
</html>
<%

Sub main()
    Dim  ssql,i,lPage,lAll,lPages,iPage
	ssql="a.id,a.tbuser,a.addtime,a.topic,a.ip,a.logid"
    select Case cmd
        Case 0
            sql="select "&ssql&" from oblog_trackback a,oblog_log b where b.userid="&oblog.l_uid&" and a.logid=b.logid order by a.ID desc"
            sGuide = sGuide & "��������ͨ��"
    End select
    Set rs = Server.CreateObject("Adodb.RecordSet")
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
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
    	<%
    	Exit Sub
    End If
    i=0
    iPage=10
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
					<li><a href="#" onclick="if (chk_idBatch(myform,'ɾ��ѡ�е�������?')==true) {document.myform.action.value='del'; document.myform.submit();}">ɾ������</a></li>
					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="TrackBackTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2">����</td>
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
					<form name="myform" method="Post" action="user_tb.asp?action=del" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
					<table id="TrackBack" class="TableList" cellpadding="0">
						<%
						'Do while not rs.EOF
						Do While Not rs.Eof And i < rs.PageSize
						i = i + 1
						%>
						<tr id="u<%=rs("ID")%>" onclick="chk_iddiv('<%=rs("id")%>')">
							<td class="t1" title="���ѡ��">
								<input name='id' type='checkbox' id="c<%=rs("ID")%>" value='<%=rs("ID")%>' onclick="chk_iddiv('<%=rs("ID")%>')" />
							</td>
							<td class="t2">
								<a href="showtb.asp?id=<%=rs("logid")%>#t<%=rs("id")%>" target="_blank"><%=oblog.filt_html(rs("topic"))%></a><br />
								<!--ʱ��-->
								<div class="time"><%=rs("tbuser")%>&nbsp;&nbsp;<%=rs("addtime")%></div>
							</td>
							<td class="t3">
								<%=rs("ip")%>
							</td>
							<td class="t4">
								<a href="user_tb.asp?action=del&id=<%=rs("ID")%>" onclick="return confirm ('ȷ��ɾ��������ͨ�棿');"><span class="red">ɾ��</span></a>
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
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/72.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
    <%
End Sub

Sub deltb()
    Dim blog, rstComment
    If id = "" Then
        oblog.adderrstr ("������ָ��Ҫɾ�������ã�")
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        id = FilterIDs(id)
        Dim n, i
        n = Split(id, ",")
        For i = 0 To UBound(n)
            delonetb (n(i))
        Next
    Else
        delonetb (id)
    End If
    oblog.ShowMsg "ɾ�����óɹ�!", ""
End Sub

Sub delonetb(id)
    Dim  logid
    id = CLng(id)
    Dim uid, mainid
    sql = "select a.logid from oblog_trackback a,oblog_log b where a.ID=" & CLng (id) & " and b.userid=" & oblog.l_uId&" and a.logid=b.logid"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then
        logid = rs(0)
        rs.Delete
        rs.Close
        '���¼���������Ŀ
        oblog.Execute ("update [oblog_log] set trackbacknum=trackbacknum-1 where logid=" & logid)
    Else
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("������ɾ��Ȩ�ޣ�")
        oblog.showusererr
        Exit Sub
    End If
End Sub
%>