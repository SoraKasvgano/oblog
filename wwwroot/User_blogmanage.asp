<!--#include file="user_top.asp"-->
<%
'��ȡ�û���־ר������
Dim rsSubject,sClass
Dim allsub,substr
Set rsSubject=Server.CreateObject("Adodb.Recordset")
rsSubject.Open "select subjectid,subjectname From oblog_subject Where userid=" & oblog.l_uid & " And subjecttype=" &t,conn,1,3
If rsSubject.Eof Then
'	sClass= "����û�����������־����"
	sClass="<select name=subjectid1 id=subjectid1 disabled>"
	sClass=sClass & "<option value=0>û����־ר��</option>"
	sClass=sClass & "</select>"
Else
	sClass="<select name=subjectid1 id=subjectid1>"
	Do While Not rsSubject.Eof
		substr=substr&"<option value="&rsSubject(0)&">"&rsSubject(1)&"</option>"
		allsub=allsub&rsSubject(0)&"!!??(("&rsSubject(1)&"##))=="
		sClass=sClass & "<option value=" & rsSubject("subjectid") & ">" & rsSubject("subjectname") & "</option>"
		rsSubject.movenext
	Loop
	sClass=sClass & "</select>"
End If

%>
<script src="oBlogStyle/move.js" type="text/javascript"></script>
<script language=javascript>
<!--
function moveaction(){
	var chkclassid1 = document.getElementById('chkclassid1');
	var chksubjectid1 = document.getElementById('chksubjectid1');
	document.myform.action.value="move";
	if (chkclassid1.checked) document.getElementById('chkclassid').value=1;
	if (chksubjectid1.checked) document.getElementById('chksubjectid').value=1;
	if ((chksubjectid1.checked||chkclassid1.checked)==0){
		alert("��ѡ��ת������");
		return false;
		}
	document.myform.classid.value=document.getElementById('classid1').value;
	document.myform.subjectid.value=document.getElementById('subjectid1').value;
	document.myform.submit();
	ShowHide("2",null);
	}
function initialize()
{
	var a = new xWin("1",300,150,292,40,"���·���","<p>���·���</p><p><input name='Submit' type='image' src='oBlogStyle/UserAdmin/4/btu_ok.png' value='ȷ���޸�' onclick='ShowHide(\"1\",null)' style='cursor:pointer;' /></p>");
	var b = new xWin("2",200,80,386,40,"�ƶ�ר��","<p> <input type='Checkbox' name='chkclassid1' id='chkclassid1' value=1>ϵͳ����:<select name='classid1' id='classid1'><%=oblog.show_Postclass(0)%></select><br><br><input type='Checkbox' name='chksubjectid1' id='chksubjectid1' value=1>Ŀ��ר��:<%=sClass%><br/><br/><input name='Submit' type='button'  value='�ƶ�' onclick='moveaction();' style='cursor:pointer;' />&nbsp;&nbsp;<input name='Submit' type='button' value='�ر�' onclick='ShowHide(\"2\",null);'/></p>");
	ShowHide("1","none");//���ش���
	ShowHide("2","none");//���ش���
}
window.onload = initialize;
//-->
</script>
<%
Dim rs,sql,id, usersearch, Keyword, sField, uid, action
Dim selectsub, wsql, truedel ,isdraft
truedel=False
isdraft=False
Keyword = Trim(Request("keyword"))
If Keyword <> "" Then Keyword = oblog.filt_badstr(Keyword)
sField = Trim(Request("Field"))
usersearch = Trim(Request("usersearch"))
selectsub = Trim(Request("selectsub"))
action = Trim(Request("action"))
id = Request("id")
If id<>"" And Instr(id,",")<=0 Then id=CLng(id)
uid = CLng(Request("uid"))
G_P_FileName = "user_blogmanage.asp?t=" & t
If usersearch = "" Then
    usersearch = 0
Else
    usersearch = CLng(usersearch)
    G_P_FileName = "user_blogmanage.asp?usersearch=" & usersearch & "&t=" & t
End If
'Request("truedel")ɾ����ƪ��־
if usersearch=6 Or LCase(Request("truedel"))="true" Then truedel=True
if usersearch=5 then isdraft=True
If selectsub <> "" Then
    selectsub = CLng(selectsub)
    G_P_FileName = "user_blogmanage.asp?usersearch=10&selectsub=" & selectsub & "&t=" & t
Else
    selectsub = 0
End If
If Keyword <> "" Then
    G_P_FileName="user_blogmanage.asp?usersearch=10&keyword="&Keyword&"&Field="&sField & "&t="  & t
End If

G_P_FileName =G_P_FileName & "&page="
If Request("page") <> "" Then G_P_This = Int(Request("page")) Else G_P_This = 1
wsql=" and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" )"

select Case LCase(action)
    Case "del"
        Call delblog
    Case "move"
        Call moveblog
    Case "updatelog"
        Call updatelog
    Case "downlog"
        Call BackUp
	Case "delall"
		Call delallblog
    Case Else
        Call main
End select
Set rs = Nothing
%>
	<tfoot>
		<tr>
			<td>
				<form name="form1" class="Search" id="ArchivesSearch" action="user_blogmanage.asp" method="get">
					<input type="hidden" name="t" value="<%=t%>">
					���ٲ��ң�&nbsp;
					<select size=1 name="usersearch" onChange="javascript:submit()">
						<option value="10" selected>��ѡ���������</option>
						<option value="0">�г�����<%=tName%></option>
						<option value="1">δͨ����˵�<%=tName%></option>
						<option value="2">��ͨ����˵�<%=tName%></option>
						<option value="3">�Ƽ�<%=tName%></option>
						<option value="5">�ݸ���</option>
					</select>
					&nbsp;&nbsp;��ר����ң�&nbsp;
					<select name="selectsub" id="selectsub" onChange="javascript:submit()">
						<option value=''>��ѡ��ר��</option>
						<%=substr%>
						<option value=0>δ����</option>
					</select>
					&nbsp;������
					<select name="Field" id="Field">
						<option value="id"><%=tName%>ID��</option>
						<option value="topic" selected><%=tName%>����</option>
						<option value="tag">��ǩ(TAG)</option>
					</select>
					 <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30" />
					 <input type="submit" name="Submit2" id="Submit" value="����" />
			  </form>
			</td>
		</tr>
	</tfoot>
</table>
</body>
</html>
<%
Set rsSubject=Nothing

Sub main()
    Server.ScriptTimeOut = 999999999
    Dim  selectsql,i,lPage,lAll,lPages,iPage,logfile
    selectsql = "TOP 500 logid,userid,iis,commentnum,topic,author,addtime,logfile,isbest,isdraft,passcheck,subjectid,istop,ispassword,ishide,classid,authorid,isspecial"
    G_P_Guide = ""
    select Case usersearch
        Case 0
            sql="select "&selectsql&" from oblog_log where isdel=0 And logtype=" & t & " And ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&") order by istop desc,addtime desc"
            G_P_Guide = G_P_Guide & "����500ƪ" & tName
        Case 1
            sql="select "&selectsql&" from [oblog_log] where  isdel=0 And passcheck=0 And logtype=" & t & " and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" ) order by addtime desc"
            G_P_Guide = G_P_Guide & "δͨ�����" & tName
        Case 2
            sql="select "&selectsql&" from [oblog_log] where  isdel=0 And passcheck=1  And logtype=" & t & " and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" ) order by addtime desc"
            G_P_Guide = G_P_Guide & "��ͨ�����" & tName
        Case 3
            sql="select "&selectsql&" from [oblog_log] where  isdel=0 And isbest=1  And logtype=" & t & " and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" ) order by addtime desc"
            G_P_Guide = G_P_Guide & "�Ƽ�" & tName
        Case 4
            sql="select "&selectsql&" from [oblog_log] where  isdel=0 And ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" )  And logtype=" & t & " order by addtime desc"
            G_P_Guide = G_P_Guide & "�ҵ�" & tName
        Case 5
            sql="select "&selectsql&" from [oblog_log] where  isdel=0 And isdraft=1  And logtype=" & t & "  and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" ) order by addtime desc"
            G_P_Guide = G_P_Guide & "�ݸ���"
		 Case 6
            sql="select "&selectsql&" from [oblog_log] where  isdel=1  And logtype=" & t & "  and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" ) order by addtime desc"
            G_P_Guide = G_P_Guide & "����վ"
        Case 10
            If Keyword = "" Then
                sql="select "&selectsql&" from [oblog_log] where  isdel=0 And (userid="&oblog.l_uid&" or authorid="&oblog.l_uid&") and subjectid="&selectsub&"   And logtype=" & t & " order by addtime desc"
                G_P_Guide=G_P_Guide & "ר��idΪ"&selectsub&"��" & tName
            Else
                select Case sField
                Case "id"
                    If IsNumeric(Keyword) = False Then
                        oblog.adderrstr (tName & "id������������")
                        oblog.showusererr
                    Else
                        sql="select "&selectsql&" from [oblog_log] where  isdel=0 And logid =" & CLng(Keyword)&"  And logtype=" & t & " and (userid="&oblog.l_uid&" or authorid="&oblog.l_uid&")"
                        G_P_Guide = G_P_Guide & "id����<font color=red> " & CLng(Keyword) & " </font>��" & tName
                    End If
                Case "topic"
                    sql="select "&selectsql&" from [oblog_log] where  isdel=0 And topic like '%" & Keyword & "%' and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" )   And logtype=" & t & " order by addtime desc"
                    G_P_Guide = G_P_Guide & "�����к��С� <font color=red>" & Keyword & "</font> ����" & tName
                Case "tag"
                    sql="select "&selectsql&" from [oblog_log] where  isdel=0 And logTags like '%" & Keyword & "%' and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" )   And logtype=" & t & " order by addtime desc"
                    G_P_Guide = G_P_Guide & "��ǩ�к��С� <font color=red>" & Keyword & "</font> ����" & tName
                Case "content"
                    sql="select "&selectsql&" from [oblog_log] where  isdel=0 And logtext like '%" & Keyword & "%' and ( userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" )   And logtype=" & t & " order by addtime desc"
                    G_P_Guide = G_P_Guide & "�����к��С� <font color=red>" & Keyword & "</font> ����" & tName
                End select
            End If
        Case Else
            oblog.adderrstr ("����Ĳ���")
            oblog.showusererr
    End select
    Set rs = Server.CreateObject("Adodb.RecordSet")
    'Response.Write(sql)
    rs.open sql, conn, 1, 3
    lAll=Int(rs.recordcount)
    If lAll=0 Then
    	%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="purl('user_post.asp','������־')">������־</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<!-- û����ؼ�¼ -->
					<div class="msg"><%=G_P_Guide & " û����ؼ�¼"%></div>
					<!-- û����ؼ�¼ end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/72.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
		<%
    	rs.Close
    	Set rs=Nothing
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
					<li><a href="#" onclick="if (chk_idBatch(myform,'ɾ��ѡ�е�<%=tName%>��?')==true) {document.myform.action.value='del'; document.myform.submit();}"><%if truedel then Response.Write("����ɾ��") else Response.Write("��־ɾ��")%></a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'���·���ѡ�е�<%=tName%>��?')==true) {document.myform.action.value='updatelog'; document.myform.submit();}"><%if truedel then Response.Write("�ָ���־") else Response.Write("���·���")%></a></li>
					<%if not truedel then%>
					<!-- <li><a href="#" onClick="return doMenu('swin1');">�ƶ�ר��</a></li> -->
					<li><a href="#" onclick="ShowHide('2',null);return false;">�ƶ�ר��</a></li>
					<%Else%>
					<li><a href="#" onclick="chk_idAll(myform,1);if (chk_idBatch(myform,'��ջ���վ�����е�<%=tName%>��?')==true) {document.myform.action.value='delall'; document.myform.submit();}">��ջ���վ</a></li>
					<%end if%>
					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="BlogManageTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"></td>
						<td class="t3"><%=G_P_Guide%></td>
						<td class="t4">�㣯��</td>
						<td class="t5">����</td>
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
					<form name="myform" method="Post" action="user_blogmanage.asp?t=<%=t%>&usersearch=<%=usersearch%>" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
					<table id="BlogManage" class="TableList" cellpadding="0">
						<%
						Do While Not rs.Eof And i < rs.PageSize
							i=i+1
						%>
						<tr id="u<%=rs("logid")%>"  onclick="chk_iddiv('<%=rs("logid")%>')">
							<td class="t1" title="���ѡ��">
								<input name='id' type='checkbox' id="c<%=rs("logid")%>" value='<%=rs("logid")%>' onclick="chk_iddiv('<%=rs("logid")%>')"/>
							</td>
							<td class="t2">
								<%
								Dim thisSubName
								thiSsubName=getsubname(rs("subjectid"),allsub)
								If thiSsubName="����" Then
									Response.Write "<span class=""grey"">��ר��</span>"
								Else
									Response.Write thiSsubName
								End If
								%>
							</td>
							<td class="t3">
							<%If rs("logfile")<>"" And rs("isdraft")= 0 Then
								If rs("isspecial") > 0 Then
									logfile = "more.asp?id="&rs("logid")
								Else
									logfile = rs("logfile")
								End if

							%>
								<%If rs("userid")<>rs("authorid") Then %>[��ͬ׫д]<%End if%>
								<%If rs("passcheck")=0 Then Response.Write "[����]"%>
								<a href="<%=logfile%>" target="_blank" title="���⣺<%=AnsiToUnicode(oblog.filt_html(rs("topic")))%>
���ڣ�<%=FormatDateTime(rs("addtime"),0)%>
���ࣺ<%
		Response.Write oblog.GetClassName(2,0,rs("classid"))
		%>
ר�⣺<%
		thiSsubName=getsubname(rs("subjectid"),allsub)
		If thiSsubName="����" Then
			Response.Write "δ����"
		Else
			Response.Write thiSsubName
		End If
		%>
�����<%=rs("iis")%>
���ۣ�<%=rs("commentnum")%>"><%=AnsiToUnicode(oblog.filt_html(rs("topic")))%></a>
							<%Else%>
								<span class="grey" onclick="purl('user_blogmanage.asp?usersearch=5','�ݸ���')">[�ݸ�]</span>&nbsp;<a href="user_post.asp?logid=<%=rs("logid")%>"  title="���⣺<%=oblog.filt_html(rs("topic"))%>
���ڣ�<%=FormatDateTime(rs("addtime"),0)%>
���ࣺ<%=oblog.GetClassName(2,0,rs("classid"))%>
ר�⣺<%
		thiSsubName=getsubname(rs("subjectid"),allsub)
		If thiSsubName="����" Then
			Response.Write "δ����"
		Else
			Response.Write thiSsubName
		End If
		%>
�����<%=rs("iis")%>
���ۣ�<%=rs("commentnum")%>"><%=oblog.filt_html(rs("topic"))%></a>
							<%End If%>
								<%
								If rs("istop")=1 Then
								%>
									<img src="oBlogStyle/li/page_up.gif" alt="��ƪ���±�������Ϊ�̶�" align="absmiddle" />
								<%
								End If
								If rs("isbest")=1 Then
								%>
									<img src="oBlogStyle/li/page_favourites.gif" alt="��ƪ���±�ϵͳ����Ա����Ϊ����" align="absmiddle" />
								 <%
								End If
								If OB_IIF(rs("ispassword"),"")<>"" Then
								%>
									<img src="oBlogStyle/li/page_key.gif" alt="��ƪ���±�������Ϊ����" align="absmiddle" />
								 <%
								End If
								If rs("ishide")=1 Then
								%>
									<img src="oBlogStyle/li/page_user.gif" alt="��ƪ���±�������Ϊ���أ����Ժ��ѿɼ�" align="absmiddle" />
								<%
								End If
								%>
								<!--ʱ��-->
								<div class="time"><%=FormatDateTime(rs("addtime"),0)%></div>
							</td>
							<td class="t4">
								<%=rs("iis")&"/"&rs("commentnum")%>
							</td>
							<td class="t5">
								<%if truedel then%>
									<a href="user_blogmanage.asp?action=updatelog&id=<%=rs("logid")%>" onClick="return confirm('ȷ��Ҫ�ָ�����־��');"><span class="blue">�ָ�</span></a>&nbsp;
								<%else%>
								<a href="user_post.asp?logid=<%=rs("logid")%>"  title="<%=tName%><%=oblog.filt_html(rs("topic"))%>"><span class="green">�޸�</span></a>&nbsp;
								<%End if%>
								<a href="user_blogmanage.asp?action=del&id=<%=rs("logid")%>&truedel=<%=truedel%>" onClick="return confirm('ȷ��Ҫɾ������־��');"><span class="red">ɾ��</span></a>
							</td>
						</tr>
						<%
						rs.MoveNext
						Loop
						rs.Close
						Set rs = Nothing
						%>
					</table>
					<input type="hidden" name="action" id="action" value="" />
					<input type="hidden" name="subjectid" id="subjectid" value="" />
					<input type="hidden" name="chksubjectid" id="chksubjectid" value="" />
					<input type="hidden" name="chkclassid" id="chkclassid" value="" />
					<input type="hidden" name="classid" id="classid" value="" />
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/90.js" type="text/javascript"></script>
				<div id="swin1" style="display:none;position:absolute;top:34px;left:342px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td align='center' class='win_table_top'>�ƶ�ר��</td>
						</tr>
						<tr>
							<td>
								<p><label for="tb">Ŀ��ר�⣺&nbsp;<input name="Submit" type="button"  value="�� ��" title="�ƶ�ר��" onclick="moveaction();" style="cursor:pointer;" /></p>
							</td>
						</tr>
						<tr>
							<td class="win_table_end"><input type="button" onClick="return doMenu('swin2');" value=" ȡ �� " title="ȡ��" /></td>
						</tr>
					</table>
				</div>
				<div id="swin2" style="display:none;"></div>
				<div id="swin3" style="display:none;"></div>
				<div id="swin4" style="display:none;"></div>
				<div id="swin5"></div>
				<iframe id="DivShim" scrolling="no" frameborder="0" style="position:absolute;top:0px; left:0px;display:none"></iframe>
			</td>
		</tr>
	</tbody>
<%
End Sub
%>

<%
Sub delblog()
    If id = "" Then
        oblog.adderrstr ("��ָ��Ҫɾ����" & tName)
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        Dim n, i
        id = FilterIDs(id)
        n = Split(id, ",")
        For i = 0 To UBound(n)
            deloneblog (n(i))
        Next
    Else
        deloneblog (id)
    End If
	Response.Write("<script>parent.get_draft();window.location='"&oblog.comeurl&"';</script>")
	Response.Flush()
End Sub

Sub deloneblog(logid)
    logid = CLng(logid)
    Dim uid, delname, rst, fso, sid,Scores,sYear,sMonth,CID,blog
    Set rst = Server.CreateObject("adodb.recordset")
    If Not IsObject(conn) Then link_database
    rst.open "select userid,logfile,subjectid,logtype,scores,isdel,addtime,isdraft,CLASSID from oblog_log where logid="&logid&wsql,conn,1,3
    If rst.Eof Then
        rst.Close
        Set rst = Nothing
        Exit Sub
    End If
		Set blog = New class_blog
	uid = rst(0)
	sYear=Year(rst(6))
	sMonth=Month(rst(6))
	delname = Trim(rst(1))
	sid = rst(2)
	CID = RST(8)
	If rst("isdraft") = 1 Then isdraft = True
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
		Call blog.DeleteFiles(logid,"")
		rst.Delete
	End If
	rst.Close
	'--------------------------------------------
	'���¼�����
	If not truedel Then
		oblog.Execute ("Update oblog_comment Set isdel=1 where mainid=" & CLng(logid))
		If Not isdraft Then
			Call OBLOG.log_count(uid,logid,sid,CID,"-")
			Call oblog.GiveScore("",-1*Abs(oblog.CacheScores(3)),"")
		End if
	Else
		Call Tags_UserDelete(logid)
		'ɾ��DIGG
		Dim RSDIGG
		Set RSDIGG = oblog.Execute ("SELECT COUNT(did) FROM oblog_digg WHERE diggtype = -1 AND logid = " & logid)
		If Not RSDIGG.Eof Then
			oblog.GiveScore "",-1*Abs(oblog.CacheScores(22))*RSDIGG(0),uid
		End If
		oblog.Execute ("DELETE FROM oblog_userdigg WHERE logid = "&logid)
		oblog.Execute ("DELETE FROM oblog_digg WHERE logid = "&logid)
		Set RSDIGG = Nothing
		oblog.Execute ("delete from oblog_comment where mainid=" & CLng(logid))
	End If
	'ɾ������
	'--------------------------------------------
	Set rst=oblog.Execute("select Count(logid) From oblog_log Where isdel=0 and isdraft=0 And Year(addtime)=" & sYear & " And Month(addtime)=" & sMonth)
	'���������ļ��Ĵ���
	If rst(0)=0 Then
		On Error Resume Next
		fso.delete Server.Mappath(blogdir & oblog.l_udir & oblog.l_ufolder & "/calendar/" & cYear & Right("0" & sMonth,2) & ".htm" )
	End If
	blog.userid = uid
	blog.Update_Subject uid
	blog.Update_index 0
	blog.Update_newblog (uid)
	Set blog = Nothing
	Set fso = Nothing
	Set rst = Nothing
End Sub

Sub delallblog()
	Dim uid, delname, rst, fso, sid,Scores,logid,blog
	Set rst = Server.CreateObject("adodb.recordset")
	If Not IsObject(conn) Then link_database
	rst.open "select userid,logfile,subjectid,logtype,logid,isdel from oblog_log where isdel=1"&wsql,conn,1,3
	If rst.Eof Then
		rst.Close
		Set rst = Nothing
		Exit Sub
	End If
	Set blog = New class_blog
	While Not rst.eof
		uid = rst(0)
		delname = Trim(rst(1))
		sid = rst(2)
		logid=rst(4)
		'�����ļ���¼
		Call blog.DeleteFiles(logid,"")
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
		'������ɾ��
		'Response.Write(truedel)
		'Response.End()
		rst.Delete
		'--------------------------------------------
		Call Tags_UserDelete(logid)
		'ɾ��DIGG
		Dim RSDIGG
		Set RSDIGG = oblog.Execute ("SELECT COUNT(did) FROM oblog_digg WHERE diggtype = -1 AND logid = " & logid)
		If Not RSDIGG.Eof Then
			oblog.GiveScore "",-1*Abs(oblog.CacheScores(22))*RSDIGG(0),uid
		End If
		oblog.Execute ("DELETE FROM oblog_userdigg WHERE logid = "&logid)
		oblog.Execute ("DELETE FROM oblog_digg WHERE logid = "&logid)
		Set RSDIGG = Nothing
		rst.MoveNext
	Wend
	rst.Close
	'���¼�����
	oblog.Execute ("delete from oblog_comment where isdel=1 ")
	'--------------------------------------------
	blog.userid = uid
	blog.Update_Subject uid
	blog.Update_index 0
	blog.Update_newblog (uid)
	Set blog = Nothing
	Set fso = Nothing
	Set rst = Nothing
	Response.Write("<script>parent.get_draft();window.location='"&oblog.comeurl&"';</script>")
	Response.Flush()
End Sub

Sub moveblog()
    If id = "" Then
        oblog.adderrstr ("��ָ��Ҫ�ƶ���" & tName)
        oblog.showusererr
        Exit Sub
    End If
    Dim subjectid,classid,chkclass,chksubject
	Dim rs,rsSubject,ishide
    chkclass=Request("chkclassid")
    chksubject=Request("chksubjectid")
    subjectid = Trim(Request("subjectid"))
    classid = Trim(Request("classid"))
    If chksubject="1" Then
	    If subjectid = 0 Then
	        oblog.adderrstr ("��ָ��Ҫ�ƶ���Ŀ��ר��")
	        oblog.showusererr
	        Exit Sub
	    Else
	        subjectid = CLng(subjectid)
	    End If
	End If
	If chkclass="1" Then
	    If classid = 0 Then
	        oblog.adderrstr ("��ָ��Ҫ�ƶ���ϵͳ����")
	        oblog.showusererr
	        Exit Sub
	    Else
	        classid = CLng(classid)
	    End If
	End If
	'��ѯĿ��ר��ID�Ƿ�Ϊ����
	Set rsSubject = oblog.Execute ("SELECT ishide FROM oblog_subject WHERE subjectid = "&subjectid)
	If Not rsSubject.Eof Then
		If rsSubject(0) = 1 Then
			ishide = True
		Else
			ishide = False
		End If
	End If
	Set rsSubject = Nothing
	Dim blog, rs1
	Set blog = New class_blog
	blog.userid = oblog.l_uId
    If InStr(id, ",") > 0 Then
        id = FilterIDs(id)
        If chksubject="1" Then
			if not IsObject(conn) then link_database
			Set rs = Server.CreateObject("Adodb.Recordset")
			rs.Open "SELECT a.subjectid ,a.isspecial , b.ishide,a.logid FROM oblog_log a LEFT JOIN oblog_subject b ON a.subjectid = b.subjectid WHERE logid in (" & id & ")  and ( a.userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" )",conn,1,3
			While Not rs.Eof
				rs(0) = subjectid
				If OB_IIF(rs(2),0) = 0 Then
					If ishide Then
						rs(1) = rs(1) + 1
					End If
				Else
					If Not ishide Then
						rs(1) = rs(1) - 1
					End If
				End If
				rs.Update
				blog.update_log rs(3),0
				rs.MoveNext
			Wend
		End if
        If chkclass="1" Then conn.execute("Update [oblog_log] set classid="&classid&" where logid in (" & id & ")"&wsql)
    Else
         If chksubject="1" Then
			if not IsObject(conn) then link_database
			Set rs = Server.CreateObject("Adodb.Recordset")
			rs.Open "SELECT a.subjectid ,a.isspecial , b.ishide,a.logid FROM oblog_log a LEFT JOIN oblog_subject b ON a.subjectid = b.subjectid WHERE logid = "&Int (id)&" and ( a.userid="&oblog.l_uid&" or authorid="&oblog.l_uid&" )",conn,1,3
			if Not rs.Eof Then
				rs(0) = subjectid
				If rs(2) = 1 Then
					If Not ishide Then
						rs(1) = rs(1) + 1
					End If
				Else
					If ishide Then
						rs(1) = rs(1) - 1
					End If
				End If
				rs.Update
				blog.update_log rs(3),0
			End if
		End If
        If chkclass="1" Then conn.execute("Update [oblog_log] set classid="&classid&" where logid=" & CLng(id) &wsql)
    End If
	Set rs = Nothing
    If chksubject="1" Then
	    Set rs = oblog.Execute("select subjectid from oblog_subject where userid=" & oblog.l_uId & " And Subjecttype=" & t)
	    While Not rs.EOF
	        Set rs1 = oblog.Execute("select count(logid) from oblog_log where oblog_log.subjectid=" & rs(0))
	        oblog.Execute ("update oblog_subject set subjectlognum=" & rs1(0) & " where oblog_subject.subjectid=" & rs(0))
	        rs.Movenext
	    Wend
	    Set rs = Nothing
	    Set rs1 = Nothing
	    blog.Update_Subject oblog.l_uId
		blog.update_index 0
	    Set blog = Nothing
	    oblog.ShowMsg "ת����־ר��ɹ�!", ""
    End If
    If chkclass="1" Then oblog.ShowMsg "����ϵͳ����ɹ�!", ""
End Sub

Sub updatelog()
	Dim aIds,i,trs,tuid,sid,cid
	Dim log_isTrouble,isdel
	Response.Write ("<table id=""TableBody"" cellpadding=""0"">") & vbcrlf
	Response.Write ("	<tbody>") & vbcrlf
	Response.Write ("		<tr>") & vbcrlf
	Response.Write ("			<td>") & vbcrlf
	Response.Write ("				<div id=""chk_idAll"">") & vbcrlf
	Response.Write ("					<div id=""prompt"">") & vbcrlf
	Response.Write ("						<ul>") & vbcrlf
	id = FilterIds(Id)
	Dim blog, p, rs, uid
	Set blog = New class_blog
	aIds=Split(id,",")
	blog.progress_init
	p = 6
	blog.progress Int(1 / p * 100), "���ɾ�̬" & tName & "�ļ�"
	blog.progress Int(2 / p * 100), "����" & tName & "�ļ�"
	log_isTrouble = 0
	For i=0 To UBound(aIds)
		Set trs = Server.CreateObject("adodb.recordset")
		trs.open "select userid,topic,abstract,logtext,isdraft,isdel,subjectid,classid FROM oblog_log WHERE logid="&aIds(i)&wsql,conn,1,3
		If trs.eof Then
			trs.close
			Exit Sub
		Else

			tuid=CLng(trs(0))
			Dim iChk1,iChk2,iChk3
			iChk1=oblog.chk_badword(trs(1))
			iChk2=oblog.chk_badword(trs(2))
			iChk3=oblog.chk_badword(trs(3))
			If trs(4) = 1 Then isdraft = True
			If trs(5) = 1 Then isdel = True
			sid = trs(6)
			cid = trs(7)
			If iChk1=0.1 Or iChk2=0.1 Or iChk3=0.1 Then
				'��¼
				oblog.execute("Update oblog_user Set isTrouble=isTrouble+1 Where userid=" & oblog.l_uid)
				'дϵͳ��־
				Dim rstLog
				Set rstLog=Server.CreateObject("Adodb.Recordset")
				rstLog.Open "select * From oblog_syslog Where 1=0",conn,1,3
				rstLog.AddNew
				rstLog("username")=oblog.l_uname
				rstLog("addtime")=oblog.ServerDate(Now)
				rstLog("addip")=oblog.userip
				rstLog("desc")="�û�����"&oblog.l_uname & "(ID��" & oblog.l_uid & ")" & " �� " & oblog.ServerDate(Now()) & " �� " & oblog.userip & " ����һƪ���°������½�ֹ����Ĺؼ��֣����±���ֹ������:<br /><font color=red>��־���⣺" & EncodeJP(trs(1)) & "<br/>���ɹؼ��֣�" & oblog.ShowBadWord & "</font>"
				rstLog("itype")=2 '�û���־��Դ
				rstLog.Update
				rstLog.Close
				oblog.adderrstr ("�����д��ھ��Խ�ֹ�Ĺؼ���,��ע����������!")

				'�ж��Ƿ���Ҫ���
				If oblog.CacheConfig(13)<>"0" And  Trim(oblog.CacheConfig(13))<>"" Then
					Dim isRedirect
					rstLog.Open "select istrouble,lockuser From oblog_user Where userid=" & oblog.l_uid,conn,1,3
					If rstLog(0)>CInt(oblog.CacheConfig(13)) Then
						rstLog("lockuser")=1
						rstLog.Update
						rstLog.Close
						isRedirect = 1
					End If
				End If
				Set rstLog=Nothing
				If oblog.errstr <> "" Then
					If isRedirect = 1 Then
						Session ("CheckUserLogined_"&oblog.l_uName) = ""
						Oblog.CheckUserLogined()
						Response.write "							<script language=JavaScript>alert('�������������ֹ��࣬�Ѿ��������');top.location='index.asp';</script>" & vbcrlf
						Response.End
					Else
						Response.Write "							<script language=JavaScript>alert(""" & oblog.errstr & """);history.go(-1)</script>" & vbcrlf
						Response.End
					End If
				End If
			Elseif iChk1 >=1 Or iChk2>=1 Or iChk3>=1 Then
				log_isTrouble=1
			End If
			trs.update
			trs.close
			Set trs=Nothing
		End If


		oblog.execute("update oblog_log set isdraft=0,isdel=0,istrouble="&log_isTrouble&" where logid="&aIds(i)&wsql)

		If isdraft Or isdel Then
			Call oblog.GiveScore("",oblog.CacheScores(3),"")
			Call OBLOG.log_count(tuid,aIds(i),sid,CID,"+")
		End if
		set rs=oblog.execute("select userid,subjectid from oblog_log where logid="&aIds(i)&wsql)
		If Not rs.EOF Then
			oblog.Execute("update oblog_comment set isdel=0 where mainid=" & aIds(i))
			blog.userid = rs(0)
			blog.Update_log aIds(i), 1
			blog.Update_calendar (aIds(i))
		Else
			Set rs = Nothing
			oblog.adderrstr ("�޲���Ȩ��!")
			oblog.showusererr
		End If
	Next
	'����ٽ�����ҳ/����ĸ���
	blog.progress Int(3 / p * 100), "������ҳ"
	blog.Update_index 0
	blog.progress Int(4 / p * 100), "����" & tName & "�����б�"
	blog.Update_Subject oblog.l_uid
	blog.progress Int(5 / p * 100), "������" & tName & "�б�"
	blog.Update_newblog oblog.l_uid
	blog.progress Int(6 / p * 100), "����" & tName & "���"
	Set rs = Nothing
	Response.Clear
	Response.Write("							<script>parent.get_draft();</script>") & vbcrlf
	Response.Write ("							<li><a href='user_blogmanage.asp'>������־����</a></li>") & vbcrlf
	Response.Write ("						</ul>") & vbcrlf
	Response.Write ("					</div>") & vbcrlf
	Response.Write ("				</div>") & vbcrlf
	Response.Write ("			</td>") & vbcrlf
	Response.Write ("		</tr>") & vbcrlf
	Response.Write ("	</tbody>") & vbcrlf
End Sub

Sub BackUp()
%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="purl('user_blogmanage.asp','��־����')">��־����</a></li>
					<li><a href="#" onclick="purl('user_blogmanage.asp?usersearch=5','�ݸ���')">�ݸ���</a></li>
					<li><a href="#" onclick="purl('user_blogmanage.asp?usersearch=6','����վ')">����վ</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUp" class="FieldsetForm">
						<legend>��ѡ����־���ݵ���ֹ���ڣ�</legend>
						<form name="oblogform" method="post" action="user_logzip.asp?t=<%=t%>">
							<ul>
								<li>��ʼ���ڣ�<%oblog.type_dateselect dateadd("m",-1,date),1%></li>
								<li>�������ڣ�<%oblog.type_dateselect date(),2%></li>
								<li>���ݸ�ʽ��<label><input name="filetype" type="radio" value="txt" checked>TXT���ı�</label>&nbsp;<label><input type="radio" name="filetype" value="htm">HTM��ҳ</label>&nbsp;<label><input type="radio" name="filetype" value="xml">XML</label></li>
								<li><input type="submit" name="addsubmit" id="Submit" value="��������"  /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/72.js" type="text/javascript"></script>
<%end Sub%>