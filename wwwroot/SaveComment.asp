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
if oblog.ChkPost()=false then Response.write("��������ⲿ�ύ!"):Response.End()
oblog.chk_commenttime
dim logid,rs,username,password,blog,isguest,comment,mainuserid,commenttopic,sCheck,checkuserlogined
logid=CLng(Request("logid"))
username=oblog.filt_badstr(Trim(Request.form("username")))
comment=Trim(Request.form("oblog_edittext"))
commenttopic=Trim(Request.Form("commenttopic"))
if username="" or oblog.strLength(username)>20 then oblog.adderrstr("���ֲ���Ϊ���Ҳ��ܴ���20���ַ���")
if oblog.chk_badword(username)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
if comment="" or oblog.strLength(comment)>Int(oblog.CacheConfig(35)) then oblog.adderrstr("�ظ����ݲ���Ϊ���Ҳ��ܴ���"&oblog.CacheConfig(35)&"Ӣ���ַ���")
if oblog.chk_badword(comment)>0 then oblog.adderrstr("�ظ������к���ϵͳ��������ַ���")
if commenttopic="" or oblog.strLength(commenttopic)>200 then oblog.adderrstr("�ظ����ⲻ��Ϊ���Ҳ��ܴ���200Ӣ���ַ���")
if oblog.chk_badword(commenttopic)>0 then oblog.adderrstr("�ظ������к���ϵͳ��������ַ���")
if oblog.chk_badword(Request.Form("homepage"))>0 then oblog.adderrstr("��ҳ��ַ�к���ϵͳ��������ַ���")
If ChkCommentTag(comment)=False Then
	sCheck=antiSpam("1")
	If sCheck<>"" Then oblog.adderrstr(sCheck)
Else
	 oblog.adderrstr("�ظ������к���ϵͳ��������ַ�")
End If
if oblog.errstr<>"" then oblog.ShowMsg Replace(oblog.errstr,"_","\n"),"back"
comment= EncodeJP(oblog.filtpath(oblog.filt_badword(comment)))
if oblog.errstr<>"" then oblog.showerr:Response.End()

password=Trim(Request.form("password"))

if oblog.CacheConfig(27)=0 or password<>"" then
		password=md5(password)
		oblog.ob_chklogin username,password,0	
end if
checkuserlogined = oblog.checkuserlogined
If checkuserlogined Then 
isguest=0
Else
	isguest=1
	Response.Cookies(cookies_name).Expires = Date + 999
	Response.Cookies(cookies_name)("username")=username
End If 
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
		oblog.ShowMsg "��֤���������ˢ����֤����������룡","back"
		Response.End()
	Else
		if not oblog.codepass then oblog.ShowMsg "��֤���������ˢ����֤����������룡","back":Response.End()
	End If
end if
if checkuserlogined And password = "" Then username = oblog.l_uname:isguest=0
if checkuserlogined And username <> oblog.l_uname Then isguest=1
If Not checkuserlogined Then isguest=1
'Process...
set blog=new class_blog
'���ӶԼ�����־�Ĵ�����ֹͨ��URL���ӻ������ʽ����־���лظ�
set rs=oblog.execute("select userid,ispassword,ishide,isencomment from oblog_log where logid="&logid)
if rs.eof then Response.Write("��������"):set rs=nothing:Response.End()
If rs("isencomment")<>"1"   Then Response.Write("����־������ظ�"):set rs=nothing:Response.End()
If Request.Cookies(cookies_name)("logpw_"&logid)<>rs("ispassword")  Then
	Response.Write("����Ĳ���!")
	Set rs=nothing
	Response.End()
End If
mainuserid=rs(0)
set rs=Server.CreateObject("adodb.recordset")
rs.open "select top 1 * from oblog_comment",conn,2,2
rs.addnew
rs("mainid")=logid
rs("userid")=mainuserid
rs("comment_user")=EncodeJP(username)
rs("commenttopic")=EncodeJP(oblog.InterceptStr(oblog.filt_badword(commenttopic),250))
rs("comment")=comment
rs("homepage")=oblog.InterceptStr(oblog.filt_badword(Request.Form("homepage")),250)
rs("addtime")=oblog.ServerDate(now())
rs("addip")=oblog.userip
rs("isguest")=isguest
rs("ubbedit")=1
rs("istate")=oblog.CacheConfig(50)
rs.update
rs.close
set rs=Nothing
Dim scores
If oblog.CacheConfig(50) = 0 Then
	scores=0
Else
	scores=oblog.CacheScores(6)
End if
oblog.execute("update oblog_user set lastcomment='" & oblog.ServerDate(Now()) &"' where userid="&mainuserid)
Response.Cookies(cookies_name)("LastComment") = oblog.ServerDate(Now())
If oblog.CacheConfig(50)=0 Then
	oblog.ShowMsg "���۳ɹ�����˺�ɼ�", ""
	Response.End
Else
	oblog.execute("update oblog_log set commentnum=commentnum+1 where logid="&logid)
	oblog.execute("update oblog_user set comment_count=comment_count+1,scores=scores+" & scores&" where userid="&mainuserid)
	oblog.execute("update oblog_setup set comment_count=comment_count+1")
	blog.userid=mainuserid
	'blog.update_comment(mainuserid)
	Server.ScriptTimeOut=99999
	blog.update_log logid,3
	blog.update_comment mainuserid
End If
'Call blog.CreateFunctionPage
if Request("t")="1" then
	Response.Redirect("more.asp?id="&logid)
Else
	Response.Redirect(blog.gourl)
end if
set blog=nothing
%>