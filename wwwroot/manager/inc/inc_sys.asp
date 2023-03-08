<!--#include file="../../conn.asp"-->
<!--#include file="../../inc/class_sys.asp"-->
<!--#include file="../../inc/md5.asp"-->
<%
Dim oblog,rstRole
set oblog=new class_sys
oblog.start
chk_sysadmin
'页面级Role
If Session("roleid")<>"" Then
	If Session("roleid") = -1 Then Response.Redirect "m_login.asp" :Response.End
	Set rstRole=oblog.Execute("select * From oblog_roles Where roleid=" & Cint(Session("roleid")))
	If Not oblog.CheckAdmin(1) Or Not Session("roleid") = "0" Then session("r_classes1")=rstRole("r_classes1")
End If
dim m_name
dim Action,FoundErr,ErrMSg,sGuide
G_P_PerMax=20
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
sub chk_sysadmin()
	dim m_password,sql,rs
	If Session("adminname")<>"" Then
		m_name=oblog.filt_badstr(session("adminname"))
		m_password=oblog.filt_badstr(session("adminpassword"))
	Else
		m_name=oblog.filt_badstr(session("m_name"))
		m_password=oblog.filt_badstr(session("m_pwd"))
	End If
	if m_name="" then
		Response.redirect "m_login.asp"
		exit sub
	end if
	sql="select id,roleid from oblog_admin where username='" & m_name & "' and password='"&m_password&"'"
	If Not IsObject(conn) Then link_database
	set rs=conn.execute(sql)
	if rs.eof then
		set rs=nothing
		Response.redirect "m_login.asp"
		exit sub
	Else
		
		If rs("roleid") = -1 Then
			set rs=nothing
			Response.redirect "m_login.asp"
			exit Sub
		Else
			If Session("roleid")="" Then
				Session("roleid") = rs("roleid")
			End if
		End if
	end if
	rs.close
	set rs=nothing

end sub

Sub WriteErrMsg()
	Dim strErr
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "<tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "<tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	Response.write strErr
End Sub

'权限校验
Function CheckAccess(field)
	On Error Resume Next
	CheckAccess=false
	If oblog.CheckAdmin(1) Or Session("roleid") = "0" Then
			CheckAccess=true
			Exit Function
	Else
		If Session("roleid")<>"" Then
			If Session("roleid") = -1 Then
				CheckAccess=False
				Exit Function
			End if
			If Not rstRole.Eof Then
				If rstRole(field)=1 Then
					CheckAccess=true
					Exit Function
				End If
			End If
		End If
	End If
	CheckAccess=False
End Function

Function CheckGoUser(ByVal usergroup)
	On Error Resume Next
	CheckGoUser=False
	If oblog.CheckAdmin(1) Or Session("roleid") = "0" Then
			CheckGoUser=true
			Exit Function
	Else
		If Session("roleid")<>"" Then
			If Session("roleid") = -1 Then
				CheckGoUser=False
				Exit Function
			End if
			If Not rstRole.Eof Then
				If rstRole("r_groups")="" Or IsNull(rstRole("r_groups"))  Then
					CheckGoUser=true
					Exit Function
				Else
					Dim arrayList,i
					arrayList = Split(rstRole("r_groups"),",")
					For i = 0 To UBound(arrayList)
						If usergroup = Int(arrayList(i)) Then
								CheckGoUser=true
								Exit Function
						End if
					Next
				End If
			End If
		End If
	End If
	CheckGoUser=False
End Function

Function CheckDisplay (stype)
	Dim ArrTemp,i
	select Case stype
	Case 1
		ReDim ArrTemp(4)
		ArrTemp(0) = "r_words"
		ArrTemp(1) = "r_IP"
		ArrTemp(2) = "r_site_news"
		ArrTemp(3) = "r_user_news"
		ArrTemp(4) = "r_site_count"
	Case 2
		ReDim ArrTemp(6)
		ArrTemp(0) = "r_user_blog"
		ArrTemp(1) = "r_user_rblog"
		ArrTemp(2) = "r_user_cmt"
		ArrTemp(3) = "r_user_msg"
		ArrTemp(4) = "r_user_tag"
		ArrTemp(5) = "r_album_comment"
		ArrTemp(6) = "r_user_digg"
	Case 3
		ReDim ArrTemp(1)
		ArrTemp(0) = "r_group_user"
		ArrTemp(1) = "r_group_blog"
	Case 4
		ReDim ArrTemp(1)
		ArrTemp(0) = "r_user_upfiles"
		ArrTemp(1) = "r_list_upfiles"
	Case 5
		ReDim ArrTemp(5)
		ArrTemp(0) = "r_user_all"
		ArrTemp(1) = "r_blogstar"
		ArrTemp(2) = "r_user_name"
		ArrTemp(3) = "r_user_admin"
		ArrTemp(4) = "r_user_add"
		ArrTemp(5) = "r_user_group"
	Case 6
		ReDim ArrTemp(1)
		ArrTemp(0) = "r_skin_sys"
		ArrTemp(1) = "r_skin_user"
	Case Else

	End select
	If oblog.CheckAdmin(1) Or Session("roleid") = "0" Then
		CheckDisplay= ""
		Exit Function
	End if
	If Not rstRole.Eof Then
		For i = 0 To UBound (ArrTemp)
			If rstRole(ArrTemp(i))=1 Then
				CheckDisplay= ""
				Exit Function
			End If
		Next
	End if
	CheckDisplay = "style='display:none;'"
End Function

Sub WriteSysLog(ByVal sContents,ByVal Strings)
	Dim sIP,rs
	sIP=oblog.userIp
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "select * From oblog_syslog Where 1=0",conn,1,3
	rs.AddNew
	rs("username")=OB_IIF(session("m_name"),session("adminname"))
	rs("addtime")=oblog.ServerDate(Now)
	rs("addip")=sIP
	rs("desc")=OB_IIF(session("m_name"),session("adminname")) & " 于 " & oblog.ServerDate(Now()) & " 自 " & sIP  & " " & sContents
	rs("QueryStrings") = Strings
	rs("itype") = 3		'内容管理员操作记录
	rs.Update
	rs.Close
	Set rs=Nothing
End Sub
%>
<script language="javascript">
function CheckSel(Voption,Value)
{
	var obj = document.getElementById(Voption);
	for (i=0;i<obj.length;i++){
		if (obj.options[i].value==Value){
		obj.options[i].selected=true;
		break;
		}
	}
}
function CheckAll(form)
{
  var v = form.chkAll.checked;
  for (var i=0;i < form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkOne")
       e.checked = v;
    }
}
function chang_size(num,objname)
{
	var obj=document.getElementById(objname)
	if (parseInt(obj.rows)+num>=3) {
		obj.rows = parseInt(obj.rows) + num;
	}
	if (num>0)
	{
		obj.width="90%";
	}
}
</script>
<style type="text/css">
#showpage{
	CLEAR: both;
	text-align: center;
	width: 100%;
}
</style>