<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/inc_antispam.asp"-->
<!--#include file="inc/md5.asp"-->
<%
'------------------------------------------------
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
If Application(cache_name_user&"_systemenmod")<>"" Then
	Dim enStr
	enStr=Application(cache_name_user&"_systemenmod")
	enStr=Split(enStr,",")
	If enStr(0)="1" Then	Response.write("ϵͳ��ʱ��ֹ����!"):Response.End()
End If
if oblog.ChkPost()=false then Response.write("��������ⲿ�ύ"):Response.End()
oblog.chk_commenttime
dim userid,rs,username,password,blog,isguest,message,mainuserid,messagetopic,homepage,ishide,sCheck,checkuserlogined
userid=CLng(Request("userid"))
username=Trim(Request.form("username"))
message=Trim(Request.form("oblog_edittext"))
messagetopic=Trim(Request.Form("commenttopic"))
ishide=Trim(Request.Form("ishide"))
homepage=oblog.InterceptStr(Request.Form("homepage"),250)
if username="" or oblog.strLength(username)>20 then oblog.adderrstr("���ֲ���Ϊ���Ҳ��ܴ���20���ַ���")
if oblog.chk_badword(username)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
if message="" or oblog.strLength(message)>Int(oblog.CacheConfig(35)) then oblog.adderrstr("�������ݲ���Ϊ���Ҳ��ܴ���"&oblog.CacheConfig(35)&"Ӣ���ַ�)��")
if oblog.chk_badword(message)>0 then oblog.adderrstr("���������к���ϵͳ��������ַ���")
if messagetopic="" or oblog.strLength(messagetopic)>200 then oblog.adderrstr("���Ա��ⲻ��Ϊ���Ҳ��ܴ���200Ӣ���ַ�)��")
if oblog.chk_badword(messagetopic)>0 then oblog.adderrstr("���Ա����к���ϵͳ��������ַ���")
if oblog.chk_badword(homepage)>0 then oblog.adderrstr("��ҳ�����к���ϵͳ��������ַ���")
'If C_SEditor_HTML=0 Then message=RemoveUBB(RemoveHtml(message))
If ChkCommentTag(message)=false Then
	sCheck=antiSpam("2")
	If sCheck<>"" Then oblog.adderrstr(sCheck)
Else
	 oblog.adderrstr("�ظ������к���ϵͳ��������ַ�")
End If
if oblog.errstr<>"" then oblog.ShowMsg Replace(oblog.errstr,"_","\n"),"back"
message=EncodeJP(oblog.filtpath(oblog.filt_badword(message)))
isguest=1
password=Trim(Request.form("password"))
if ishide<>"" then ishide=cint(ishide) else ishide=0
if oblog.CacheConfig(27)=0 or password<>"" then
		password=md5(password)
		oblog.ob_chklogin username,password,0	
end if
checkuserlogined = oblog.checkuserlogined
If checkuserlogined Then isguest=0


if oblog.CacheConfig(27)=0 Then
	If Not checkuserlogined then
		oblog.ShowMsg "��Ҫ��¼����ܷ�������","back"
		Response.End
	Else
		If oblog.l_ulevel="6" Then
			oblog.ShowMsg "��û��ͨ������Ա��ˣ����ܷ�������","back"
			Response.End()
		End If
	End If
End If

if oblog.CacheConfig(30)=1 Then
	If  Request("CodeStr")="" then
		oblog.ShowMsg "��֤���������������","back"
		Response.End()
	Else
		if not oblog.codepass then oblog.ShowMsg "��֤���������������","":Response.End()
	End If
end if
if checkuserlogined And password = "" Then username = oblog.l_uname:isguest=0
'Process...
set blog=new class_blog
if not IsObject(conn) then link_database
set rs=Server.CreateObject("adodb.recordset")
rs.open "select top 1 * from oblog_message",conn,2,2
rs.addnew
rs("userid")=userid
rs("message_user")=EncodeJP(oblog.filt_badword(username))
rs("messagetopic")=EncodeJP(oblog.InterceptStr(oblog.filt_badword(messagetopic),250))
rs("message")=message
rs("homepage")=homepage
rs("addtime")=oblog.ServerDate(now())
rs("addip")=oblog.userip
rs("isguest")=isguest
rs("ishide")=ishide
rs("istate")=oblog.CacheConfig(50)
rs.update
rs.close
set rs=Nothing
Dim scores
If oblog.CacheConfig(50) = 0 Then
	scores=0
Else
	scores=oblog.CacheScores(5)
End if
oblog.execute("update oblog_user set lastmessage='" & oblog.ServerDate(Now()) &"' where userid="&userid)
Response.Cookies(cookies_name)("LastComment") = oblog.ServerDate(now())
If oblog.CacheConfig(50)=0 Then
	Set blog=Nothing
	oblog.ShowMsg "���Գɹ�����˺�ɼ�", ""
Else
	oblog.execute("update oblog_user set message_count=message_count+1 ,scores=scores+" & scores&" where userid="&userid)
	oblog.execute("update oblog_setup set message_count=message_count+1")
	blog.userid=userid
	blog.update_message 3
	blog.update_newmessage userid
	Dim GoUrl
	GoUrl=blog.gourl
	Set blog=Nothing
	Response.Redirect GoUrl
End if
%>