<!--#include file="user_top.asp"-->
<body  style="overflow:hidden;" class="user_iframe">
<%
If oblog.l_Group(17,0)=0 Then
	oblog.AddErrStr ("��Ŀǰ�����ĵȼ�������ʹ��ͨѶ¼����")
    oblog.showUserErr
    Response.End
End if
Dim action,sGuide
Dim rs,rsSubject,addrId,Ids,targetSubjectid,allsub
Dim sName,sProvince,sCity,sSex,sBirthday,sMail,sUrl,sQQ,sMsn,sMob,sSubjectId
Dim sHomeAddr,sHomeTel,sOfficeTel,sOfficeAddr,sMemo
Action = LCase(Trim(Request("action")))
Set rs=Server.CreateObject("Adodb.Recordset")
addrId=Request("id")
If addrId<>"" And InStr(addrid,",")<=0 Then addrId=CLng(addrId)
'����һ��������Ϣ
Set rsSubject=Server.CreateObject("Adodb.Recordset")
rsSubject.Open "select * From oblog_subject Where userid=" & oblog.l_uid & " And subjecttype=2",conn,1,3
select Case Action
	Case "save"
		Call Save
	Case "add","edit"
		Call EditForm
	Case "del"
		'ɾ��
		If addrId<>"" Then
			addrId=FilterIds(addrid)
			conn.Execute("Delete From oblog_address Where addrid In (" & addrid & ") And userid=" & oblog.l_uid)
		End If
		Response.Redirect "user_address.asp"
	Case "bdel"
		'����ɾ��
		Ids=FilterIds(addrid)
		If Ids<>"" Then
			conn.Execute("Delete From oblog_address Where addrid in(" & Ids & ") And userid=" & oblog.l_uid)
		End If
		Response.Redirect "user_address.asp"
	Case "bmove"
		Ids=FilterIds(addrid)
		targetSubjectid=clng(Request("subject"))
		If Ids<>"" Then 		'������ת��
			conn.Execute("Update oblog_address Set subjectid=" &targetSubjectid  &" Where addrid in(" & Ids & ") And userid=" & oblog.l_uid)
		End If
		Response.Redirect "user_address.asp"
	Case Else
		Call List
End select
Set rs=Nothing
%>
<div id="user_page_search">
	<form name="form2" action="user_address.asp?cmd=11" method="post">
	&nbsp;&nbsp;&nbsp;&nbsp;
	�����ѯ:<%
		If rsSubject.RecordCount>0 Then
			rsSubject.MoveFirst
	    	If rsSubject.Eof Then
	    		%>
	    		δ������
	    		<%
	    	Else
	    		%>
	    		<select name="subjectid"  onChange="javascript:submit()">
	    		<%
	    		Do While Not rsSubject.Eof
					allsub=allsub&rsSubject(0)&"!!??(("&rsSubject(1)&"##))=="
	    			%>
	    			<option value="<%=rsSubject("subjectid")%>"><%=rsSubject("subjectname")%></option>
	    			<%
	    			rsSubject.MoveNext
	    		Loop
	    		%>
	    		</select>
	    	<%
	    	End If
    	End If
    	%>
	�ؼ���:<input type="text" value="" name="keyword" size=30 maxlength=30>
	<input type="submit" value="��ѯ">
 	</form>
<%
rsSubject.Close
Set rsSubject=Nothing
%>
</div>
</body>
</html>

<%
Sub EditForm()
  If addrId<>"" Then
	Set rs=oblog.Execute("select * From oblog_address Where userid=" & oblog.l_uid & " And  addrid=" & ProtectSQL(addrId))
	If rs.Eof Then
		ErrMsg="�������Ϣ���"
	Else
		sName=rs("Name")
		sProvince=rs("Province")
		sCity=rs("City")
		sSex=rs("Sex")
		sBirthday=rs("Birthday")
		sMail=rs("email")
		sUrl=rs("Url")
		sQQ=rs("QQ")
		sMsn=rs("Msn")
		sMob=rs("Mobile")
		sSubjectId=rs("SubjectId")
		sHomeAddr=rs("HomeAddr")
		sHomeTel=rs("HomeTel")
		sOfficeTel=rs("OfficeTel")
		sOfficeAddr=rs("OfficeAddr")
		sMemo=rs("Memo")
	End If
	rs.Close
End If
%>
<script language=javascript>

function VerifySubmit()
{
    if (document.oblogform.name.value.length==0){
    	alert("��ϵ������������д");
    	document.oblogform.name.focus();
    	return false;
    	}
    if (document.oblogform.email.value.length==0){
    	alert("�����ʼ�������д");
    	document.oblogform.email.focus();
    	return false;
    	}
    	return true;
}
</script>
<body scroll="no" style="overflow-x:hidden;background:#fff">
<ul id="user_page_top">
	<%If addrid="" Then Response.Write "������ϵ��" Else Response.Write "ά����ϵ����Ϣ" End If%>
	&nbsp;&nbsp;&nbsp;&nbsp;[<a href="#" onclick="history.back()">������ϵ���б�</a>]
	&nbsp;&nbsp;&nbsp;&nbsp;[<a href="user_subject.asp?t=2"">��ϵ�˷���ά��</a>]
</ul>
<div id="user_setting_content">
	<div id="cnt">
    	<div id="dTab11" class="Box">
    <form action="user_address.asp?action=save&id=<%=addrid%>" method="post" name="oblogform" id="oblogform" onSubmit="return VerifySubmit()">
	<table  class="dTab13_body" align="center" border="0" cellpadding="0" cellspacing="1">
	  <tr>
		<td class="dTab13_body_td">�û�����</td>
		<td colspan="3">
   	<%
    	If rsSubject.Eof Then
    		%>
    		��Ŀǰ��û���趨ͨѶ¼���飬�����Լ�����ӻ���<a href="user_subject.asp?t=2">�趨��������</a>
    		<%
    	Else
    		%>
    		<select name="subjectid">
    		<%
    		Do While Not rsSubject.Eof
    			%>
    			<option value="<%=rsSubject("subjectid")%>" <%If rsSubject("subjectid")=sSubjectId Then Response.Write "checked" End If%>><%=rsSubject("subjectname")%></option>
    			<%
    			rsSubject.MoveNext
    		Loop
    		%>
    		</select>
    	<%
    	End If
    	%>
		</td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">�ա���</td>
		<td><input name="name" type=text size="20" maxlength="250" value="<%=sName%>"></td>
		<td class="dTab13_body_td">�ԡ���</td>
		<td><input name="sex" type=radio value=1 <%If sSex="1" Then Response.Write "checked"%>>��<input name="sex" type=radio value=2 <%If sSex="2" Then Response.Write "checked"%>>Ů</td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">������</td>
		<td><input name="birthday" type=text  size="20" maxlength="12" value="<%=sBirthday%>">&nbsp;(��:20050601)</td>
		<td class="dTab13_body_td">���ڵ�</td>
		<td><%=oblog.type_city(sProvince,sCity)%></td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">Email</td>
		<td><input name="email" type=text  size="20" maxlength="200" value="<%=sMail%>"><font color="#FF0000"> *</font></td>
		<td class="dTab13_body_td">�ֻ�����</td>
		<td><input name="mob" type=text  size="20" maxlength="20" value="<%=sMob%>"></td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">QQ</td>
		<td><input name="qq" type=text  size="20" maxlength="200" value="<%=sQQ%>"></td>
		<td class="dTab13_body_td">MSN</td>
		<td><input name="msn" type=text  size="20" maxlength="200" value="<%=sMsn%>"></td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">��ͥסַ</td>
		<td><input name="homeaddr" type=text  size="50" maxlength="200" value="<%=sHomeAddr%>"></td>
		<td class="dTab13_body_td">��ͥ�绰</td>
		<td><input name="hometel" type=text  size="20" maxlength=20 value="<%=sHomeTel%>"></td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">��˾��ַ</td>
		<td><input name="officeaddr" type=text  size="50" maxlength="200" value="<%=sOfficeAddr%>"></td>
		<td class="dTab13_body_td">��˾�绰</td>
		<td><input name="officetel" type=text  size="20" maxlength=200 value="<%=sOfficeTel%>"></td>
	  </tr>
	  <tr>
		<td colspan="4" align="center"><input type="submit" value=" �� �� " /> <input type="reset" value=" �� �� " /></td>
	  </tr>
	</table>
    </form>
   </div>
  </div>
 </div>

<%
End Sub

Sub Save()
    If addrId<>"" Then addrId=CLng(addrId)
    If oblog.ChkPost() = False Then
        oblog.AddErrStr ("ϵͳ��������ⲿ�ύ��")
        oblog.showUserErr
        Exit Sub
    End If
    'Get
    sName=Request.Form("Name")
	sProvince=Request.Form("Province")
	sCity=Request.Form("City")
	sSex=Request.Form("Sex")
	sBirthday=Request.Form("Birthday")
	sMail=Request.Form("email")
	sUrl=Request.Form("Url")
	sQQ=Request.Form("QQ")
	sMsn=Request.Form("Msn")
	sMob=Request.Form("Mob")
	sSubjectId=Request.Form("SubjectId")
	sHomeAddr=Request.Form("HomeAddr")
	sHomeTel=Request.Form("HomeTel")
	sOfficeTel=Request.Form("OfficeTel")
	sOfficeAddr=Request.Form("OfficeAddr")
	sMemo=Request.Form("Memo")
    'Check
    If sName = "" Or oblog.strLength(sName) > 50 Then oblog.AddErrStr ("��������Ϊ���Ҳ��ܴ���50���ַ�����")
    If sMail = "" Or oblog.strLength(sMail) > 50 Then oblog.AddErrStr ("Email����Ϊ���Ҳ��ܴ���50���ַ�����")
    If oblog.strLength(sMemo)>1000 Then oblog.AddErrStr ("��ע���ݲ��ܴ���1000���ַ�����")
    If oblog.chk_badword(sProvince) > 0 Then oblog.AddErrStr ("����ѡ���к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sCity) > 0 Then oblog.AddErrStr ("����ѡ���к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sMail) > 0 Then oblog.AddErrStr ("Email�к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sUrl) > 0 Then oblog.AddErrStr ("��ҳ�к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sMob) > 0 Then oblog.AddErrStr ("�ֻ������к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sQQ) > 0 Then oblog.AddErrStr ("QQ�����к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sMsn) > 0 Then oblog.AddErrStr ("MSN�к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sHomeAddr) > 0 Then oblog.AddErrStr ("��ͥסַ�к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sHomeTel) > 0 Then oblog.AddErrStr ("��ͥ�绰�к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sOfficeTel) > 0 Then oblog.AddErrStr ("�����к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sOfficeAddr) > 0 Then oblog.AddErrStr ("��˾��ַ�к���ϵͳ���������Ĺؼ��֣�")
    If oblog.chk_badword(sMemo) > 0 Then oblog.AddErrStr ("��˾�绰�к���ϵͳ���������Ĺؼ��֣�")

    If oblog.ErrStr <> "" Then oblog.showUserErr
    If addrId<>"" Then
    	rs.Open "select * From oblog_address Where addrId=" & addrId & " And userid=" & oblog.l_uid,conn,1,3
    	If rs.Eof Then
    		rs.Close
    		Set rs=Nothing
    		oblog.AddErrStr ("Ŀ�����ݲ����ڣ��뷵�����²�����")
        oblog.showUserErr
    	End If
  	Else
      rs.Open "select * From oblog_address Where 1=0",conn,1,3
      rs.AddNew
   	End If
    '��ʼд�����
    rs("name") =  EncodeJP(oblog.filt_badword(sName))
    rs("email") = oblog.filt_badword(sMail)
    rs("url") = sUrl
    rs("subjectid") = sSubjectId
    rs("birthday") = sBirthday
    rs("qq") = OB_IIF(sQQ,"-")
    rs("msn")=OB_IIF(sMsn,"-")
    rs("sex")=sSex
    rs("Province")=sProvince
    rs("city")=sCity
    rs("mobile")=OB_IIF(sMob,"-")
    rs("homeaddr")=OB_IIF(sHomeAddr,"-")
    rs("officeaddr")=OB_IIF(sOfficeAddr,"-")
    rs("hometel")=OB_IIF(sHomeTel,"-")
    rs("officetel")=OB_IIF(sOfficeTel,"-")
    rs("userid")=oblog.l_uid
    rs("addtime") = oblog.ServerDate(Now)
    rs.Update
    rs.Close
    Response.Redirect "user_address.asp"
End Sub

Sub List()

	Dim Sql,i,lPage,lAll,lPages,iPage,Subjectid,keyword,cmd
	Subjectid=Request("Subjectid")
	keyword=Request("keyword")
	If Keyword <> "" Then Keyword = oblog.filt_badstr(Keyword)
	cmd=LCase(Request("cmd"))
	select Case cmd
		Case "11"
			If keyword<>"" Then
				Sql="select addrid,subjectid,name,email,url From oblog_address Where userid=" & oblog.l_uid & " And (name like '%" & keyword &"%' Or email like '%" & keyword &"%' Or qq like '%" & keyword &"%')   Order By subjectid,addtime Desc"
			Else
				If Subjectid<>"" Then
					Subjectid=CLng(Subjectid)
					Sql="select addrid,subjectid,name,email,url From oblog_address Where userid=" & oblog.l_uid & " And subjectid=" & subjectid &" Order By subjectid,addtime Desc"
				End If
			End If
		Case Else
			Sql="select addrid,subjectid,name,email,url From oblog_address Where userid=" & oblog.l_uid & " Order By subjectid,addtime Desc"
	End select
	rs.Open Sql,conn,1,3
	lAll=Int(rs.recordcount)
    If lAll=0 Then
    	rs.Close
    	Set rs=Nothing
    	%>
    	<div id="user_page_content">
		   <div id="content_li">
		   	<ul class="content_li_conten">
		   		<li class="t1"></li>
		   		<li class="t3">&nbsp;</li>
		   	</ul>
		   	<ul class="content_li_conten">
		   		<li class="t1"></li>
		   		<li class="t3"><%=sGuide & " û����ؼ�¼" %>,<a href="user_address.asp?action=add">����һ����ϵ��</a></li></ul>
		  	</div>
		  </div>
    	<%
    	Exit Sub
    End If
    iPage=12
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
<body style="background:#fff">
<style type="text/css">
<!--
	.content_li_top .t1 {width:50px;text-align:center;}
	.content_li_top .t2 {width:110px;}
	.content_li_top .t3 {width:125px;}
	.content_li_top .t4 {width:160px;text-align:left;}
	.content_li_top .t5 {width:100px;}
	#content_li .content_li_conten .t1 {width:40px;text-align:center;}
	#content_li .content_li_conten .t2 {width:110px;color:#999;}
	#content_li .content_li_conten .t3 {width:130px;}
	#content_li .content_li_conten .t4 {width:160px;text-align:left;}
	#content_li .content_li_conten .t5 {width:100px;}
-->
</style>
	<ul id="user_page_top">
		<li id="p7"><a href="#" onclick="chk_idAll(myform,1)">ȫ��ѡ��</a></li>
		<li id="p8"><a href="#" onclick="chk_idAll(myform,0)">ȫ��ȡ��</a></li>
		<li id="p4"><a href="#" onclick="if (chk_idBatch(myform,'ɾ��ѡ�е���ϵ����?')==true) { document.myform.submit();}">ɾ����ϵ��</a></li>
		<li>&nbsp;&nbsp;&nbsp;&nbsp;</li>
		<li id="p1"><a href="user_address.asp?action=add">������ϵ��</a></li>
		<li id="p1"><a href="user_subject.asp?t=2"">����ά��</a></li>
	</ul>
	<div id="showpage">
	  <%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
	</div>
	<div id="user_page_content">
		<ul class="content_li_top">
			<li class="t1">ѡ��</li>
			<li class="t2">����</li>
			<li class="t3">����</li>
			<li class="t4">Email</li>
			<li class="t5">����</li>
		</ul>
  		 <div id="content_li">
			<form name="myform" method="Post" action="user_address.asp?action=del" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
          <%
          Do while not rs.EOF
          	i = i + 1%>
          	<ul class="content_li_conten" id="u<%=rs("addrid")%>">
		    <li class="t1"><input name='id' type='checkbox'  id="c<%=cstr(rs("addrid"))%>" value='<%=cstr(rs("addrid"))%>' onclick="chk_id('<%=rs("addrid")%>')" /></li>
		    <li class="t2">
			<%=getsubname(rs("subjectid"),allsub)%></li>
		    <li class="t3" onclick="chk_iddiv('<%=rs("addrid")%>')"><%=rs("name")%></li>
		    <li class="t4" onclick="chk_iddiv('<%=rs("addrid")%>')"><%=rs("email")%></li>
		    <li class="t5">
				<a href="user_address.asp?action=edit&id=<%=rs("addrid")%>">�鿴ά��</a>&nbsp;
				<a href="user_address.asp?action=del&id=<%=rs("addrid")%>" onClick="return confirm('ȷ��Ҫɾ������ϵ����Ϣ��');">ɾ��</a>&nbsp;
			</li>
			</ul>
<%
    If i >= iPage Then Exit Do
    rs.Movenext
Loop
%>
</form>
</div>
</div>
</body>
<%
End Sub
%>