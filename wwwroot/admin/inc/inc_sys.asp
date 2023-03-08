<!--#include file="../../conn.asp"-->
<!--#include file="../../inc/class_sys.asp"-->
<!--#include file="../../inc/inc_valid.asp"-->
<!--#include file="../../inc/md5.asp"-->
<!--#include file="../../inc/inc_control.asp"-->
<%
dim oblog
set oblog=new class_sys
oblog.start
chk_sysadmin
dim admin_name
G_P_PerMax=20
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
sub chk_sysadmin()

	dim admin_password,sql,rs
	admin_name=oblog.filt_badstr(session("adminname"))
	admin_password=oblog.filt_badstr(session("adminpassword"))
	if admin_name="" then
		Response.redirect "admin_login.asp"
		exit sub
	end if
	sql="select id,roleid from oblog_admin where username='" & admin_name & "' and password='"&admin_password&"'"
	If Not IsObject(conn) Then link_database
	set rs=conn.execute(sql)
	if rs.bof and rs.eof then
		set rs=nothing
		Response.redirect "admin_login.asp"
		exit Sub
	Else
		If rs("roleid") <>0 Then
			set rs=nothing
			Response.redirect "admin_login.asp"
			exit Sub
		End if
	end if
	rs.close
	set rs=nothing
end sub
Function CheckSafePath(byval strMode)
	Dim strPathFrom,strPathSelf,arrFrom,arrSelf,i
	CheckSafePath=False
	If oBlog.ChkPost=False Then Exit Function
	strPathFrom = Replace(LCase(CStr(Request.ServerVariables("HTTP_REFERER"))),"http://","")
    strPathSelf = Replace(LCase(CStr(Request.ServerVariables("PATH_INFO"))),"http://","")
    If strPathFrom="" Then Exit Function
    If strPathSelf="" Then Exit Function
    arrFrom=Split(strPathFrom,"/")
    arrSelf=Split(strPathSelf,"/")
    For i=0 To UBound(arrFrom)
    	'Response.Write "arrFrom("&i&")="& arrFrom(i) & "<BR/>"
	Next
	For i=0 To UBound(arrSelf)
    	'Response.Write "arrSelf("&i&")="& arrSelf(i) & "<BR/>"
	Next
    select Case strMode
    	Case "0"
			 For i = 1 To (UBound(arrSelf)-1)
	            If arrFrom(i)=arrSelf(i) And Left(arrFrom(UBound(arrfrom)),Len(arrSelf(UBound(arrself))))=arrSelf(UBound(arrself))Then CheckSafePath=True
			 Next
	End select
End Function
Sub WriteErrMsg()
	Dim strErr
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "<tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "<tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	Response.write strErr
End Sub
Sub EventLog(ByVal sDesc,ByVal Strings)'写日志
	Dim sIP,rs
	sIP=oblog.userIp
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "select * From oblog_syslog Where 1=0",conn,1,3
	rs.AddNew
	rs("username")=session("adminname")
	rs("addtime")=oblog.ServerDate(Now)
	rs("addip")=sIP
	rs("desc")=session("adminname") & " 于 " & oblog.ServerDate(Now()) & " 自 " & sIP  & " " & sDesc
	rs("QueryStrings") = Strings
	rs("itype")=1 '系统管理员操作记录
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
<script language="javascript" src="../inc/div.js"></script>
<style type="text/css">
#showpage{
	CLEAR: both;
	text-align: center;
	width: 100%;
}
</style>