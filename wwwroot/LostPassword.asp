<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<!--#include file="API/Class_API.asp"-->
<%
'------------------------------------------------
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
If is_ot_user=1 Then
	If Not IsObject(conn) Then link_database
	Response.Redirect(ot_lostpasswordurl)
	Response.End()
End If
If oblog.ChkPost() = False Then
	oblog.adderrstr ("ϵͳ��������ⲿ�ύ��")
End If
If oblog.cacheConfig(84) = "0" Then
	oblog.adderrstr ("ϵͳ��ֹ�û�ʹ��ȡ�����빦�ܣ�")
End If
if oblog.errstr<>"" then oblog.showerr
dim action,show_getpassword
action=Int(Request("action"))
call sysshow()
G_P_Show =  Replace (G_P_Show,"$show_title_list$",oblog.cacheConfig(2) & "--ȡ������")
show_getpassword="��ǰλ�ã�<a href='index.asp'>��ҳ</a>���һ�����<hr noshade>"
select case action
	case 1
	call sub_getpassword_1()
	case 2
	call sub_getpassword_2()
	case 3
	call sub_getpassword_3()
	case else
	call sub_getpassword_0()
end select


G_P_Show=Replace(G_P_Show,"$show_list$",show_getpassword)
Response.Write G_P_Show&oblog.site_bottom

dim pass_username,daan

sub sub_getpassword_0()
	show_getpassword=show_getpassword&"<form name='form1' method='post' action=''>"
	show_getpassword=show_getpassword&"<TABLE width='400' border=0 cellPadding=0 cellSpacing=0 borderColor=#111111 style='BORDER-COLLAPSE: collapse'>"
	show_getpassword=show_getpassword&"<tr><td height='18' class='bian'><strong>�һ������һ��:</strong></td> </tr><tr>"
	show_getpassword=show_getpassword&"<td height='200'><div align='center'>�������û���:"
	show_getpassword=show_getpassword&"<input name='uid' type='text' id='uid' size='23' maxlength='26'>"
	show_getpassword=show_getpassword&"<br><br><input name='Submit' type='submit' id='Submit' value='��һ��'>"
	show_getpassword=show_getpassword&"<input name='action'  id='action' type='hidden' value='1'>"
	show_getpassword=show_getpassword&"</div></td></tr></table></form>"
end sub

sub sub_getpassword_1()
	dim rs
	pass_username=oblog.filt_badstr(Trim(Request("uid")))
	if pass_username="" then oblog.adderrstr("�û�������Ϊ�գ�"):oblog.showerr:rs.close:exit Sub
	set rs=oblog.execute ("select username,Question,answer,user_group from [oblog_user] where username='"&pass_username&"'")
	if rs.eof then oblog.adderrstr("���û��������ڣ�")
	if oblog.errstr<>"" then oblog.showerr:rs.close:exit Sub
	If rs("answer") = "" Or IsNull (rs("answer")) Then
		oblog.adderrstr("������ʾ��Ϊ�գ�����ϵ����Աȡ�����룡")
		oblog.showerr
		rs.close
		exit Sub
	End If
	Dim rst
	Set rst=oblog.Execute("select g_getpwd From oblog_groups Where groupid=" & rs(3))
	If rst(0) = 0 Then
		oblog.adderrstr("��ǰ�û������û��鲻����ȡ�����룬����ϵ����Ա��")
		oblog.showerr
		rs.close
		rst.Close
		exit Sub
	End If
	Set rst = Nothing
	show_getpassword=show_getpassword&"<form name='form1' method='post' action=''>"
	show_getpassword=show_getpassword&"<TABLE width='400' border=0 cellPadding=0 cellSpacing=0 borderColor=#111111 style='BORDER-COLLAPSE: collapse'>"
	show_getpassword=show_getpassword&"<tr><td height='18' ><strong>�һ�����ڶ���:</strong></td> "
	show_getpassword=show_getpassword&"</tr><tr> <td height='200'><div align='center'>�������û���:"
	show_getpassword=show_getpassword&"<input name='uid' type='text' id='uid' value='"&rs("username")&"' size='30' maxlength='26' readonly>"
	show_getpassword=show_getpassword&"<br><br>������ʾ����:"
	show_getpassword=show_getpassword&"<input name='tishi' type='text' id='tishi' value='"&oblog.filt_html(rs("Question"))&"' size='30' maxlength='26'>"
	show_getpassword=show_getpassword&"<br><br>���������:"
	show_getpassword=show_getpassword&"<input name='daan' type='text' id='daan' size='30' maxlength='26'><br /><br />"
	show_getpassword=show_getpassword&"��֤��:<input name=""codestr"" type=""text"" size=""6"" maxlength=""20"" />"&oblog.getcode
	show_getpassword=show_getpassword&"<br><br><input name='Submit' type='submit' id='Submit' value='��һ��'>"
	show_getpassword=show_getpassword&"<input name='action'  id='action' type='hidden' value='��'>"
	show_getpassword=show_getpassword&"</div></td></tr></table></form>"
	rs.close
	set rs=nothing
end sub

sub sub_getpassword_2()
	dim tishi,rs
	pass_username=oblog.filt_badstr(Trim(Request("uid")))
	daan=Trim(Request("daan"))
 	if daan="" then
		oblog.adderrstr("�Բ���������ʾ����𰸲���Ϊ�գ�")
	Else
		daan = MD5(daan)
	End if
	if not oblog.codepass then oblog.adderrstr("��֤�����")
	if oblog.errstr<>"" then oblog.showerr:exit Sub
	set rs=oblog.execute("select userid from [oblog_user] where username='"&pass_username&"' and answer='"&daan&"'")
	if rs.eof then oblog.adderrstr("������ʾ����ش���󣡣�"):oblog.showerr:exit Sub
	Session("GetCode")="true"
	show_getpassword=show_getpassword&"<TABLE width='400' border=0 cellPadding=0 cellSpacing=0  align='center' style='BORDER-COLLAPSE: collapse'>"& vbcrlf
	show_getpassword=show_getpassword&"<tr><td height='18' class='bian'><strong>���,�������趨����:</strong></td></tr><tr>"& vbcrlf
	show_getpassword=show_getpassword&"<td><table width='100%' border='0' cellpadding='0' cellspacing='0'>"& vbcrlf
	show_getpassword=show_getpassword&"<tr><td><FORM action='lostpassword.asp' method='post' name='regform' >"& vbcrlf
	show_getpassword=show_getpassword&"<br><br><table width='60%' border='0' align='center' cellpadding='0' cellspacing='0'>"& vbcrlf
	show_getpassword=show_getpassword&"<tr><td><table  border='0' cellspacing='0' cellpadding='5'>"& vbcrlf
	show_getpassword=show_getpassword&"<tr> <td width='37%'><div align='right'>"& vbcrlf
	show_getpassword=show_getpassword&"<p>�����룺</p></div></td><td colspan='2'><input name='new_pass' type='password' id='new_pass'></td>"& vbcrlf
	show_getpassword=show_getpassword&"</tr><tr><td><div align='right'>��֤���룺</div></td>"& vbcrlf
	show_getpassword=show_getpassword&"<td colspan='2'><input name='new_pass2' type='password' id='new_pass2'></td>"& vbcrlf
	show_getpassword=show_getpassword&"</tr><tr><td><div align='right'> </div></td>"& vbcrlf
	show_getpassword=show_getpassword&"<td width='17%'><input type='submit' name='Submit' value='ȷ��'></td>"& vbcrlf
	show_getpassword=show_getpassword&"<td width='46%'><input type='reset' name='Submit2' value='ȡ��'></td>"& vbcrlf
	show_getpassword=show_getpassword&"</tr></table><input name='action'  id='action' type='hidden' value='3'><input name='uid'  id='uid' type='hidden' value='"&pass_username&"'><input name='daan'  id='daan' type='hidden' value='"&daan&"'></td></tr></table></form><br><div align='center'> </div></td>"& vbcrlf
	show_getpassword=show_getpassword&"</tr></table></td></tr></table>"& vbcrlf
	rs.close
	set rs=nothing
end sub

sub sub_getpassword_3()
	If Session("GetCode")<>"true" Then Exit Sub
	dim password,repassword
	Dim rs
	pass_username=oblog.filt_badstr(Trim(Request("uid")))
	daan=oblog.filt_badstr(Trim(Request("daan")))
	password=Trim(Request("new_pass"))
	repassword=Trim(Request("new_pass2"))
	if password="" or oblog.strLength(password)>14 or oblog.strLength(password)<4 then oblog.adderrstr("���벻��Ϊ��(���ܴ���14С��4)��")
	if repassword="" then oblog.adderrstr("�ظ����벻��Ϊ�գ�")
	if password<>repassword then oblog.adderrstr("�����������벻ͬ��")
	if daan="" then oblog.adderrstr("������ʾ�𰸲���Ϊ�գ�")
	if oblog.errstr<>"" then oblog.showerr:exit Sub
    If Not IsObject(conn) Then link_database
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select answer,password,TruePassWord FROM [oblog_user] WHERE username='" & pass_username & "'", conn, 1, 3
	If Not rs.EOF Then
		If rs("answer") <> daan Or rs("answer") = "" Or IsNull(rs("answer")) Then
			Exit Sub
		End If
		If API_Enable Then
			Dim blogAPI,j,strUrl
			Set blogAPI = New DPO_API_OBLOG
			blogAPI.LoadXmlFile True
			blogAPI.UserName=pass_username
			blogAPI.PassWord=password
			Call blogAPI.ProcessMultiPing("update")
			Set blogAPI=Nothing
			For j=0 To UBound(aUrls)
				strUrl=Lcase(aUrls(j))
				If Left(strUrl,7)="http://" Then
					Response.write("<script src="""&strUrl&"?syskey="&MD5(pass_username&oblog_Key)&"&username="&pass_username&"&password="&MD5(password)&"&savecookie=0""></script>")
				End If
			Next
		End If
		Dim TruePassWord
		TruePassWord = RndPassword(16)
		rs("password") = md5(password)
		rs("TruePassWord") = TruePassWord
		rs.update
	End If
	rs.Close
	Set rs = Nothing
	oblog.savecookie pass_username,TruePassWord,0
	Session("GetCode") = Empty
	show_getpassword="��ǰλ�ã�<a href='index.asp'>��ҳ</a>���޸�����ɹ�<hr noshade>"
	show_getpassword=show_getpassword&"<strong>���������Ѿ��޸ĳɹ���</strong><br>"
	show_getpassword=show_getpassword&"<a href='index.asp'>���������ҳ��</a><br>"
	show_getpassword=show_getpassword&"5����Զ���������̨��"
	show_getpassword=show_getpassword&"<script language=JavaScript>"
	show_getpassword=show_getpassword&"setTimeout(""window.location='user_index.asp'"",5000);"
	show_getpassword=show_getpassword&"</script>"
end sub

%>