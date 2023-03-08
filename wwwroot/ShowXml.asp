<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/Cls_XmlDoc.asp"-->
<%
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
Dim user_path,XmlPath
Dim Show,user_group,teamID,calendar,userid,blogname
user_path = Trim(Request("user_path"))
user_group = Request("user_group")
teamID = Request("teamid")
calendar = Request("calendar")
userid = Request("userid")
blogname = Request("blogname")
Dim xmlDoc
Set xmlDoc = New Cls_XmlDoc
On Error Resume Next
'用户页面信息
If user_path <> "" Then
	XmlPath = blogdir&user_path&"/user.config"
	If xmlDoc.LoadXml (XmlPath) Then
		Echo xmlDoc.SelectXmlNode ("comment",1)
		Echo xmlDoc.SelectXmlNode ("mygroups",1)
		Echo xmlDoc.SelectXmlNode ("newblog",1)
		Echo xmlDoc.SelectXmlNode ("newmessage",1)
		Echo xmlDoc.SelectXmlNode ("search",1)
		Echo xmlDoc.SelectXmlNode ("subject",1)
		If OBLOG.CacheConfig(81) = "1" Then Echo xmlDoc.SelectXmlNode ("aobomusic",1)
		Echo xmlDoc.SelectXmlNode ("links",1)
		Echo xmlDoc.SelectXmlNode ("myfriend",1)
		Echo xmlDoc.SelectXmlNode ("blogname",1)
		Echo xmlDoc.SelectXmlNode ("info",1)
		Echo xmlDoc.SelectXmlNode ("placard",1)
	End if
End if
'用户页面广告
If user_group <> "" Then
	Dim rst
	Set rst=oblog.Execute("select g_ad_sys From oblog_groups Where groupid=" & CLng(user_group) )
	If rst(0) = 1 Then
		XmlPath = blogdir&oblog.CacheConfig(80)&"/GG.config"
		If xmlDoc.LoadXml (XmlPath) Then
			Echo xmlDoc.SelectXmlNode ("ad_usertop",1)
			Echo xmlDoc.SelectXmlNode ("ad_usercomment",1)
			Echo xmlDoc.SelectXmlNode ("ad_userbot",1)
			Echo xmlDoc.SelectXmlNode ("ad_userlinks",1)
			'兼容旧广告
			Echo xmlDoc.SelectXmlNode ("gg_usertop",1)
			Echo xmlDoc.SelectXmlNode ("gg_usercomment",1)
			Echo xmlDoc.SelectXmlNode ("gg_userbot",1)
			Echo xmlDoc.SelectXmlNode ("gg_userlinks",1)
		End If
	End If
	rst.Close
	Set rst = Nothing
'群组页面广告
ElseIf teamID <> "" Then
	XmlPath = blogdir&oblog.CacheConfig(80)&"/GG.config"
	If xmlDoc.LoadXml (XmlPath) Then
		Echo xmlDoc.SelectXmlNode ("gg_teamtop",1)
		Echo xmlDoc.SelectXmlNode ("gg_teamcomment",1)
		Echo xmlDoc.SelectXmlNode ("gg_teambot",1)
		Echo xmlDoc.SelectXmlNode ("gg_teamlinks",1)
		'兼容旧广告
		Echo xmlDoc.SelectXmlNode ("ad_teamtop",1)
		Echo xmlDoc.SelectXmlNode ("ad_teamcomment",1)
		Echo xmlDoc.SelectXmlNode ("ad_teambot",1)
		Echo xmlDoc.SelectXmlNode ("ad_teamlinks",1)
	End If
End If
Sub Echo(sStr)
	Response.Write sStr
	Response.Flush
End Sub
Set XmlDoc = Nothing
'Response.Write oblog.htm2js_div ("<a href=""http://www.oblog.cn/rss/?rss="&oblog.CacheConfig(3) &blogdir&user_path&"/rss2.xml"" target=""_blank""><img src=""http://www.oblog.cn/xml.jpg"" border=""0"" /></a>","txml")
Response.Write oblog.htm2js_div ("<a href="""&oblog.CacheConfig(3)&"user_url.asp?action=add&mainuserid="&userid&"&sTitle="&blogname&"&sUrl="&blogdir&user_path&"/rss2.xml"" target=""_blank""><img src=""http://www.oblog.cn/xml.jpg"" border=""0"" /></a>","txml")
Set OBLOG = Nothing
%>