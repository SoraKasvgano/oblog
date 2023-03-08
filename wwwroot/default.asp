<!--#include file="inc/inc_syssite.asp"-->
<%
Dim From,gourl,turl
From = LCase(Request.ServerVariables("HTTP_HOST"))
gourl= LCase(oblog.CacheConfig(3))
turl = Replace (gourl,"http://","")
turl = Left(turl, InStrRev(turl, "/")-1)
If From  = turl Then
	Response.Redirect(gourl&"index.asp")'此处为网站首页地址
Else
	Response.Write( "<frameset><frame src="""&gourl&"blog.asp?domain="&from&"""></frameset>")
End If
Set oblog= Nothing
%>