<!--#include file="user_top.asp"-->
<body  style="overflow:hidden;" class="user_iframe">
<%
If oblog.l_Group(17,0)=0 Then
	oblog.AddErrStr ("您目前所属的等级不允许使用通讯录功能")
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
'缓存一个分类信息
Set rsSubject=Server.CreateObject("Adodb.Recordset")
rsSubject.Open "select * From oblog_subject Where userid=" & oblog.l_uid & " And subjecttype=2",conn,1,3
select Case Action
	Case "save"
		Call Save
	Case "add","edit"
		Call EditForm
	Case "del"
		'删除
		If addrId<>"" Then
			addrId=FilterIds(addrid)
			conn.Execute("Delete From oblog_address Where addrid In (" & addrid & ") And userid=" & oblog.l_uid)
		End If
		Response.Redirect "user_address.asp"
	Case "bdel"
		'批量删除
		Ids=FilterIds(addrid)
		If Ids<>"" Then
			conn.Execute("Delete From oblog_address Where addrid in(" & Ids & ") And userid=" & oblog.l_uid)
		End If
		Response.Redirect "user_address.asp"
	Case "bmove"
		Ids=FilterIds(addrid)
		targetSubjectid=clng(Request("subject"))
		If Ids<>"" Then 		'批量组转移
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
	按组查询:<%
		If rsSubject.RecordCount>0 Then
			rsSubject.MoveFirst
	    	If rsSubject.Eof Then
	    		%>
	    		未定义组
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
	关键字:<input type="text" value="" name="keyword" size=30 maxlength=30>
	<input type="submit" value="查询">
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
		ErrMsg="错误的信息编号"
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
    	alert("联系人姓名必须填写");
    	document.oblogform.name.focus();
    	return false;
    	}
    if (document.oblogform.email.value.length==0){
    	alert("电子邮件必须填写");
    	document.oblogform.email.focus();
    	return false;
    	}
    	return true;
}
</script>
<body scroll="no" style="overflow-x:hidden;background:#fff">
<ul id="user_page_top">
	<%If addrid="" Then Response.Write "新增联系人" Else Response.Write "维护联系人信息" End If%>
	&nbsp;&nbsp;&nbsp;&nbsp;[<a href="#" onclick="history.back()">返回联系人列表</a>]
	&nbsp;&nbsp;&nbsp;&nbsp;[<a href="user_subject.asp?t=2"">联系人分类维护</a>]
</ul>
<div id="user_setting_content">
	<div id="cnt">
    	<div id="dTab11" class="Box">
    <form action="user_address.asp?action=save&id=<%=addrid%>" method="post" name="oblogform" id="oblogform" onSubmit="return VerifySubmit()">
	<table  class="dTab13_body" align="center" border="0" cellpadding="0" cellspacing="1">
	  <tr>
		<td class="dTab13_body_td">用户分组</td>
		<td colspan="3">
   	<%
    	If rsSubject.Eof Then
    		%>
    		您目前还没有设定通讯录分组，您可以继续添加或者<a href="user_subject.asp?t=2">设定分组后添加</a>
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
		<td class="dTab13_body_td">姓　名</td>
		<td><input name="name" type=text size="20" maxlength="250" value="<%=sName%>"></td>
		<td class="dTab13_body_td">性　别</td>
		<td><input name="sex" type=radio value=1 <%If sSex="1" Then Response.Write "checked"%>>男<input name="sex" type=radio value=2 <%If sSex="2" Then Response.Write "checked"%>>女</td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">生　日</td>
		<td><input name="birthday" type=text  size="20" maxlength="12" value="<%=sBirthday%>">&nbsp;(如:20050601)</td>
		<td class="dTab13_body_td">所在地</td>
		<td><%=oblog.type_city(sProvince,sCity)%></td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">Email</td>
		<td><input name="email" type=text  size="20" maxlength="200" value="<%=sMail%>"><font color="#FF0000"> *</font></td>
		<td class="dTab13_body_td">手机号码</td>
		<td><input name="mob" type=text  size="20" maxlength="20" value="<%=sMob%>"></td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">QQ</td>
		<td><input name="qq" type=text  size="20" maxlength="200" value="<%=sQQ%>"></td>
		<td class="dTab13_body_td">MSN</td>
		<td><input name="msn" type=text  size="20" maxlength="200" value="<%=sMsn%>"></td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">家庭住址</td>
		<td><input name="homeaddr" type=text  size="50" maxlength="200" value="<%=sHomeAddr%>"></td>
		<td class="dTab13_body_td">家庭电话</td>
		<td><input name="hometel" type=text  size="20" maxlength=20 value="<%=sHomeTel%>"></td>
	  </tr>
	  <tr>
		<td class="dTab13_body_td">公司地址</td>
		<td><input name="officeaddr" type=text  size="50" maxlength="200" value="<%=sOfficeAddr%>"></td>
		<td class="dTab13_body_td">公司电话</td>
		<td><input name="officetel" type=text  size="20" maxlength=200 value="<%=sOfficeTel%>"></td>
	  </tr>
	  <tr>
		<td colspan="4" align="center"><input type="submit" value=" 保 存 " /> <input type="reset" value=" 清 除 " /></td>
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
        oblog.AddErrStr ("系统不允许从外部提交！")
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
    If sName = "" Or oblog.strLength(sName) > 50 Then oblog.AddErrStr ("姓名不能为空且不能大于50个字符长度")
    If sMail = "" Or oblog.strLength(sMail) > 50 Then oblog.AddErrStr ("Email不能为空且不能大于50个字符长度")
    If oblog.strLength(sMemo)>1000 Then oblog.AddErrStr ("备注内容不能大于1000个字符长度")
    If oblog.chk_badword(sProvince) > 0 Then oblog.AddErrStr ("地区选择中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sCity) > 0 Then oblog.AddErrStr ("地区选择中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sMail) > 0 Then oblog.AddErrStr ("Email中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sUrl) > 0 Then oblog.AddErrStr ("主页中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sMob) > 0 Then oblog.AddErrStr ("手机号码中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sQQ) > 0 Then oblog.AddErrStr ("QQ号码中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sMsn) > 0 Then oblog.AddErrStr ("MSN中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sHomeAddr) > 0 Then oblog.AddErrStr ("家庭住址中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sHomeTel) > 0 Then oblog.AddErrStr ("家庭电话中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sOfficeTel) > 0 Then oblog.AddErrStr ("标题中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sOfficeAddr) > 0 Then oblog.AddErrStr ("公司地址中含有系统不允许发布的关键字！")
    If oblog.chk_badword(sMemo) > 0 Then oblog.AddErrStr ("公司电话中含有系统不允许发布的关键字！")

    If oblog.ErrStr <> "" Then oblog.showUserErr
    If addrId<>"" Then
    	rs.Open "select * From oblog_address Where addrId=" & addrId & " And userid=" & oblog.l_uid,conn,1,3
    	If rs.Eof Then
    		rs.Close
    		Set rs=Nothing
    		oblog.AddErrStr ("目标数据不存在，请返回重新操作！")
        oblog.showUserErr
    	End If
  	Else
      rs.Open "select * From oblog_address Where 1=0",conn,1,3
      rs.AddNew
   	End If
    '开始写入操作
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
		   		<li class="t3"><%=sGuide & " 没有相关纪录" %>,<a href="user_address.asp?action=add">增加一个联系人</a></li></ul>
		  	</div>
		  </div>
    	<%
    	Exit Sub
    End If
    iPage=12
	'分页
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = Int(Request("page"))
	End If

	'设置缓存大小 = 每页需显示的记录数目
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
		<li id="p7"><a href="#" onclick="chk_idAll(myform,1)">全部选择</a></li>
		<li id="p8"><a href="#" onclick="chk_idAll(myform,0)">全部取消</a></li>
		<li id="p4"><a href="#" onclick="if (chk_idBatch(myform,'删除选中的联系人吗?')==true) { document.myform.submit();}">删除联系人</a></li>
		<li>&nbsp;&nbsp;&nbsp;&nbsp;</li>
		<li id="p1"><a href="user_address.asp?action=add">增加联系人</a></li>
		<li id="p1"><a href="user_subject.asp?t=2"">分类维护</a></li>
	</ul>
	<div id="showpage">
	  <%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
	</div>
	<div id="user_page_content">
		<ul class="content_li_top">
			<li class="t1">选中</li>
			<li class="t2">分类</li>
			<li class="t3">姓名</li>
			<li class="t4">Email</li>
			<li class="t5">操作</li>
		</ul>
  		 <div id="content_li">
			<form name="myform" method="Post" action="user_address.asp?action=del" onSubmit="return confirm('确定要执行选定的操作吗？');">
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
				<a href="user_address.asp?action=edit&id=<%=rs("addrid")%>">查看维护</a>&nbsp;
				<a href="user_address.asp?action=del&id=<%=rs("addrid")%>" onClick="return confirm('确定要删除此联系人信息吗？');">删除</a>&nbsp;
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