<!--#include file="../inc/inc_syssite.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
If Not oblog.CheckUserLogined Or oblog.CacheConfig(81) = "0" Then
	Response.End
End if
Dim PassPort_e,PassPort_p,userid,PostURL,uid
Dim PlayerType
Dim blog
Dim action
Set blog = New class_blog
PostURL = oblog.CacheConfig(3)
PostURL = Replace(PostURL,"http://","")
PassPort_e = ProtectSQL (Request("PassPort_e"))
PassPort_p = ProtectSQL (Request("PassPort_p"))
PlayerType = Trim (Request("PlayerType"))
userid = Trim (Request("userid"))
action = Trim (Request("action"))
If action = "aobomusic" Then Call GoMusic()
'OB_DEBUG Request.QueryString,1
If PassPort_e <>"" And PassPort_p<>"" Then
'	Response.Cookies("aobo_PassPort").Expires = Date + 365
'	If cookies_domain <> "" Then
'		Response.Cookies("aobo_PassPort").domain = cookies_domain
'	End If
'	Response.Cookies("aobo_PassPort").Path   =   blogdir
'	Response.Cookies("aobo_PassPort")("PassPort_email") = PassPort_e
'	Response.Cookies("aobo_PassPort")("PassPort_password") = PassPort_p
	If userid = 0  Then
		'È¡Ïû°ó¶¨
		oblog.Execute ("UPDATE oblog_user SET PassPort_userid = 0,PassPort_email = NULL,PassPort_password = NULL WHERE userid = "&oblog.l_uid)
		blog.userid=oblog.l_uid
		blog.update_index 0
		Set blog = Nothing
		Response.Clear
		Response.Write "<script>top.location='"&blogurl&"user_index.asp'</script>"
		Response.End
	Else
		oblog.Execute ("UPDATE oblog_user SET PassPort_userid = "&CLng(userid)&",PassPort_email = '"&PassPort_e&"',PassPort_password = '"&PassPort_p&"' WHERE userid = "&oblog.l_uid)
	End If
	Response.Redirect "http://music.aobo.com/aobomusic.php?myserverurl="&PostURL&"&passport_e="& PassPort_e &"&passport_p="& PassPort_p
Else
	If PlayerType <>"" Then
		PlayerType = CLng(PlayerType)
	Else
		PlayerType = 0
	End if
	If PlayerType = 1 And userid <>"" Then
		oblog.Execute ("UPDATE oblog_user SET PlayerType = 1 WHERE PassPort_userid = "&CLng(userid)&" AND userid = "&oblog.l_uid)
	Else
		oblog.Execute ("UPDATE oblog_user SET PlayerType = 0 WHERE PassPort_userid = "&CLng(userid)&" AND userid = "&oblog.l_uid)
	End If
	blog.userid=oblog.l_uid
	blog.update_index 0
	Set blog = Nothing
	Response.Redirect "http://music.aobo.com/aobomusic.php?w=player&myserverurl="&PostURL
End If

Sub GoMusic()
	Dim passport_email,passport_password,AOBOstr
	Dim rsAOBO
	Set rsAOBO = oblog.Execute ("SELECT userid,passport_email,passport_password FROM oblog_user WHERE userid = "&oblog.l_uid)
	If rsAOBO.Eof Then
		passport_email = ""
		passport_password = ""
	Else
		passport_email = rsAOBO(1)
		passport_password = rsAOBO(2)
	End If
	Response.Redirect "http://music.aobo.com/aobomusic.php?myserverurl="&PostURL&"&passport_e="& passport_email &"&passport_p="& passport_password
End Sub
Set oblog = Nothing
%>