<!--#include file="user_top.asp"-->
<%
'----------------------------------------
'Oblog4�����ļ�
'Ӧ��������ģ���е��û����ֻ֧࣬��һ����
'��־/���/ͨѶ¼/��ժ
'----------------------------------------
%>
<script type="text/javascript">
	function addLoadEvent(func) {if (typeof wpOnload!='function'){wpOnload=func;}else{ var oldonload=wpOnload;wpOnload=function(){oldonload();func();}}}
</script>
<script src="inc/dbx-admin-key.js" type="text/javascript"></script>
<script src="inc/dbx.compressed.js" type="text/javascript"></script>
<%
Dim rs, sql, blog
Dim id, action
action = Trim(request("action"))
id = CLng(request("id"))
Select Case action
    Case "addclass"
		Call addclass
    Case "del"
		Call delclass
    Case "modify"
		Call modifyclass
    Case "savemodi"
		Call savemodify
    Case "order"
		Call order
	Case "update"
		Call update_log
    Case Else
		Call main
End Select
Set rs = Nothing
%>
<script type="text/javascript">if(typeof wpOnload=='function')wpOnload();</script>
</body>
</html>
<%
Sub addclass()
	Dim subjectname, rs, ordernum,ishide
    subjectname = Trim(request.Form("subjectname"))
    ishide = Trim(request.Form("ishide"))
    If subjectname = "" Or oblog.strLength(subjectname) > 50 Then oblog.adderrstr ("����������Ϊ���Ҳ��ܴ���50�ַ�)��")
    If oblog.chk_badword(subjectname) > 0 Then oblog.adderrstr ("�������к���ϵͳ��������ַ���")
	If oblog.errstr<>"" Then oblog.showusererr:Exit Sub
    Set rs = oblog.execute("select max(ordernum) from oblog_subject where userid=" & oblog.l_uid & " And SubjectType=" & CLng(t))
    If Not IsNull(rs(0)) Then
        ordernum = rs(0) + 1
    Else
        ordernum = 1
    End If
    Set rs = server.CreateObject("adodb.recordset")
    rs.open "select top 1 * from [oblog_subject] Where SubjectType=" & t, conn, 1, 3
    rs.addnew
    rs("subjectname") = subjectname
    rs("userid") = oblog.l_uid
    rs("ordernum") = ordernum
    rs("subjectType") = t
	If ishide = "on" Then RS("ishide") = 1 Else rs("ishide") = 0
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ShowMsg "��ӷ���ɹ�!", "user_subject.asp?t=" & t
End Sub

Sub delclass()
    Dim id
    id = CLng(request.QueryString("id"))
    oblog.execute("delete  from [oblog_subject] where subjectid="&id&" and userid="&oblog.l_uid)
    oblog.execute("update [oblog_log] set subjectid=0 where subjectid="&id&" and userid="&oblog.l_uid)
    Call order
    oblog.ShowMsg "ɾ������ɹ�!", ""
End Sub

Sub savemodify()
    Dim subjectname,rs,ishide,goUrl
    id = CLng(id)
	goUrl = "user_subject.asp?t="&t
    subjectname = Trim(request.Form("subjectname"))
	ishide = Trim(request.Form("ishide"))
    If subjectname = "" Or oblog.strLength(subjectname) > 50 Then oblog.adderrstr ("����������Ϊ���Ҳ��ܴ���50�ַ�)��")
    If oblog.chk_badword(subjectname) > 0 Then oblog.adderrstr ("�������к���ϵͳ��������ַ���")
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = server.CreateObject("adodb.recordset")
    rs.open "select subjectname,ishide from [oblog_subject] where subjectid="&id&" and userid="&oblog.l_uid,conn,1,3
    If Not rs.EOF Then
        rs("subjectname") = subjectname
		If ishide = "on" Then
			If RS("ishide") = 0 Then
				RS("ishide") = 1
				If T = 1 Then
					Oblog.Execute ("UPDATE oblog_album SET ishide = 1 WHERE userclassid = "&id)
				ElseIf t = 0 Then
					Oblog.Execute ("UPDATE oblog_log SET isspecial = 0 WHERE isspecial IS NULL AND subjectid = "&id)
					Oblog.Execute ("UPDATE oblog_log SET isspecial = isspecial + 1 WHERE subjectid = "&id)
					'�˴��������־��̬ҳ��
					goUrl ="user_subject.asp?action=update&id="&id
				End If
			End If
		Else
			If RS("ishide") = 1 Then
				rs("ishide") = 0
				If T = 1 Then
					Oblog.Execute ("UPDATE oblog_album SET ishide = 0 WHERE userclassid = "&id)
				ElseIf t = 0 Then
					Oblog.Execute ("UPDATE oblog_log SET isspecial = 0 WHERE isspecial IS NULL AND subjectid = "&id)
					Oblog.Execute ("UPDATE oblog_log SET isspecial = isspecial - 1 WHERE isspecial >0 AND subjectid = "&id)
					'�˴��������־��̬ҳ��
					goUrl ="user_subject.asp?action=update&id="&id
				End If
			End if
		End If
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
    %>
    <script language="javascript">
    	//alert("�޸ķ������Ƴɹ�!");
    	parent.location.href="<%=goUrl%>";
  	</script>
    <%
End Sub

Sub order()
	Dim subjectid,rs,i
	subjectid = FilterIDs(Request("subjectid"))
	If subjectid = FilterIDs(Request("subjectid0")) Then
		Response.Redirect "user_subject.asp?t="&t
	End If
	subjectid = Split (subjectid,",")
    Set rs = server.CreateObject("adodb.recordset")
	For i = 0 To UBound(subjectid)
		rs.open "SELECT ordernum FROM [oblog_subject] WHERE subjectid=" & subjectid(i) & " AND userid="&oblog.l_uid ,conn,1,3
		rs(0) = i + 1
		rs.Update
		rs.Close
	Next
	Set rs = Nothing
	Response.Redirect "user_subject.asp?t="&t
End Sub

Sub main()
%>
<script language="javascript">
	/* ����ҳ�浯������ */
function doMenu1(MenuName,URL){
//	alert("���� ");
	document.getElementById("chgClass").src=URL;
	document.getElementById(MenuName).style.display = "block";
//	alert("����2 ");
//	if(document.getElementById(MenuName).style.display == "block"){
//		document.getElementById(MenuName).style.display = "none";
//	}
//	else{
//		document.getElementById(MenuName).style.display == "block";
//		}
	}
</script>
<form method="post" action="">
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onClick="return doMenu('swin1');">��ӷ���</a></li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="SubjectTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2"><%=tName%>����</td>
						<td class="t3">����</td>
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
					<table id="Subject" cellpadding="0">
						<tr>
							<td class="t1">
<%
Dim rs1
Set rs1 = oblog.execute("select * from oblog_subject where userid=" & oblog.l_uid & " And SubjectType=" & t & " order by ordernum")
While Not rs1.EOF
%>
								<div class="ordernum"><%=rs1("ordernum")%></div>
<%
rs1.movenext
Wend
Set rs1 = Nothing
%>
							</td>
							<td class="t2">
								<div class="dbx-group" id="content_li">
<%
Dim rs
Set rs = oblog.execute("select * from oblog_subject where userid=" & oblog.l_uid & " And SubjectType=" & t & " order by ordernum")
While Not rs.EOF
%>
									<div class="dbx-box">
										<span class="dbx-handle">
											<input type="hidden" name="subjectid" value=<%=rs("subjectid")%>>
											<ul>
												<li class="l1"><%
												If t= "1" Then

													Response.Write "<a href='"&blogdir&oblog.l_udir&"/"&oblog.l_ufolder&"/cmd."&f_ext&"?uid="&oblog.l_uid&"&do=album&id="&rs("subjectid")&"' target='_blank'>"&oblog.filt_html(rs("subjectname"))&"</a>"
												Else
													Response.Write "<a href='"&blogdir&oblog.l_udir&"/"&oblog.l_ufolder&"/cmd."&f_ext&"?uid="&oblog.l_uid&"&do=blogs&id="&rs("subjectid")&"' target='_blank'>"&oblog.filt_html(rs("subjectname"))&"</a>"
												End if


												%></li>
												<li class="l2"><%If t = 0 Then %><a href="user_subject.asp?action=update&id=<%=rs("subjectid")%>"><span class="red">���·���</span></a> <%End if%> <a href="javascript:void(0);" onClick="return doMenu1('swin2','user_subject.asp?action=modify&id=<%=rs("subjectid")%>&t=<%=t%>');"><span class="green">�޸�</span></a> <a href="user_subject.asp?action=del&id=<%=rs("subjectid")%>&t=<%=t%>" <%="onClick='return confirm(""ȷ��Ҫɾ���˷�����(���ɻָ�)��"");'"%>><span class="red">ɾ��</span></a></li>
											</ul>
										</span>
									</div>
									<input type="hidden" name="subjectid0" value="<%=rs("subjectid")%>">
<%
rs.movenext
Wend
Set rs = Nothing
%>
								</div>
							</td>
						</tr>
					</table>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/90.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
	<tfoot class="SubjectBottom">
		<tr>
			<td>
				<input type="hidden" name="action" value="order">
				<input type="submit" value="�������">
			</td>
		</tr>
	</tfoot>
</table>
</form>

<div id="swin1" style="display:none;position:absolute;top:34px;left:10px;z-index:100;">
	<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
		<tr>
			<td colspan='2' class='win_table_top'>���<%=tName%>����</td>
		</tr>
		<tr>
			<td colspan='2'><%If t=0 Or t=1 Or t="" Then %>
				���<%=tName%>�����ֻ���ڴ˷��෢��<%=tName%>�Ż�����ҳ��ʾ����!
				<%End If%>
			</td>
		</tr>
		<tr>
			<td class='win_table_td'><%=tName%>�������ƣ�</td>
			<td>
				<form name="form1" method="post" action="user_subject.asp?action=addclass&t=<%=t%>">
				<input name="subjectname" type="text" id="subjectname" maxlength="50" /><br />
				<label><input name="ishide" type="checkbox" id="ishide" maxlength="50" />���ط�������</label>
			</td>
		</tr>
		<tr>
			<td colspan='2' class="win_table_end"><input type="submit" value=" �� �� " title="���" />&nbsp;&nbsp;<input type="button" onClick="return doMenu('swin1');" value=" �� �� " title="�ر�" /></td>
		</tr>
	</table>
	</form>
</div>
<div id="swin2" style="display:none;position:absolute;top:50px;left:50px;z-index:100;">
<iframe class="FrmID" id="chgClass"  style="width:442px;height:154px;" src="" frameborder="0" scrolling="auto" onunload="parent.location.href='user_subject.asp?t=<%=t%>'"></iframe>
</div>
<div id="swin3"></div>
<div id="swin4"></div>
<div id="swin5"></div>
<iframe id="DivShim" scrolling="no" frameborder="0" style="position:absolute;top:0px; left:0px;display:none"></iframe>

<%
End Sub
Sub modifyclass()
    Dim oldname, rs
    id = CLng(id)
    set rs=oblog.execute("select subjectname,ishide from oblog_subject where subjectid="&id&" and userid="&oblog.l_uid)
    If Not rs.EOF Then
    oldname = oblog.filt_html(rs(0))
%>
	<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
		<tr>
			<td colspan='2' align='center' class='win_table_top'>�޸�<%=tName%>����</td>
		</tr>
		<tr>
			<td colspan='2'>����<%=tName%>����������Ҫ������ҳ�Ż�ʹ�޸���Ч!</td>
		</tr>
		<tr>
			<td class='win_table_td' rowspan=2 ><%=tName%>�������ƣ�</td>
			<td>
				<form name="form1" method="post" action="user_subject.asp?action=savemodi&id=<%=id%>&t=<%=t%>">
				<input name="subjectname" type="text" id="subjectname" maxlength="20" value="<%=oldname%>" />
			</td>
		</tr>
		<tr>
			<td>
				<label><input name="ishide" type="checkbox" id="ishide" maxlength="50" <%If rs("ishide")= 1 Then Response.Write "checked"%> />���ط�������</label>
			</td>
		</tr>
		<tr>
			<td colspan='2' class="win_table_end"> <input type="submit" value="�޸�" /> &nbsp;&nbsp;<input type="button" onClick="return parent.doMenu('swin2');" value="�ر�" title="�ر�" /></td>
		</tr>
	</table>
					</form>
<%
    Set rs = Nothing
    End If
End Sub
Sub update_log()
	Dim subjectid
	subjectid = CLng(Request("id"))
	Dim blog
	Set blog = new class_blog
	Response.Write("") & vbcrlf
	Response.Write("<div id=""prompt"">") & vbcrlf
	Response.Write("	<ul>") & vbcrlf
	blog.progress_init
	blog.Update_subjectlog oblog.l_uid,subjectid
	blog.update_index 0
	Response.Write("		<li><a href='javascript:history.go(-1)'>������һҳ</a></li>") & vbcrlf
	Response.Write("	</ul>") & vbcrlf
	Response.Write("</div>") & vbcrlf
End Sub
%>