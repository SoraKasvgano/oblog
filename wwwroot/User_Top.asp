<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="API/Class_API.asp" -->
<%
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
Dim t,tName,nowurl,cssfile,actiontype,url

url=Request.QueryString("url")
t=Request("t")
actiontype=Request("actiontype")
nowurl=oblog.NowUrl
dim uurl

G_P_PerMax=20
'退出系统
If Request("t")="logout" Then

    oblog.CheckUserLogined
    If API_Enable Then
		Dim strUrl,j
		For j=0 To UBound(aUrls)
			strUrl=Lcase(aUrls(j))
			If Left(strUrl,7)="http://" Then
				Response.write("<script language=JavaScript src="""&strUrl&"?syskey="&MD5(oblog.l_uName&oblog_Key)&"&username="&oblog.l_uName&"&password=""></script>")&vbcrlf
			End If
		Next
	End If

	If cookies_domain <> "" Then
        Response.Cookies(cookies_name).domain = cookies_domain
    End If
	Response.Cookies(cookies_name).Path   =   blogdir
	Response.Cookies(cookies_name)("username")=""
	Response.Cookies(cookies_name)("password")=""
	Response.Cookies(cookies_name)("userurl")=""
	If API_Enable Then
		Response.write "<script language=JavaScript>setTimeout(""window.location='index.asp'"",1000);</script>"
	Else
		if Request("re")="1" then
			Response.Write("<script language=JavaScript>top.location='"&oblog.comeurl&"';</script>")
		else
			Response.Write("<script language=JavaScript>top.location='index.asp';</script>")
		End If
	End If
	Session ("CheckUserLogined_"&oblog.l_uName) = ""
'	Session.Abandon
	Set oBlog=Nothing
	Response.End()
End If

select Case t
    Case 0, ""
    	t = 0
        tName = "日志"
    Case 1
        tName = "相册"
    Case 2
        tName = "通讯录"
    Case 3
        tName = "订阅"
    Case Else
        t = 0
        tName = "日志"
End select
if not oblog.checkuserlogined() then
	Response.Redirect("login.asp?fromurl="&Replace(oblog.GetUrl,"&","$"))
else
	if oblog.l_ulevel=6 then
		oblog.adderrstr("您未通过管理员审核，不能进入后台")
		oblog.showerr
	end if
end if
'服务器端控制禁止页面单独出现
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<script>
var ismusicok=0
function purl(url,title){
	parent.show_title(title);
	parent.switchTab('TabPage1','Tab3');
	parent.menu(parent.document.getElementById('Tab3'))
	parent.frames["content3"].location = url;
}


function pmusicurl(){
parent.switchTab('TabPage1','Tab2');
parent.menu(parent.document.getElementById('Tab2'))
if(ismusicok == 0){parent.frames["content2"].location = 'api/api_aobo.asp?action=aobomusic';ismusicok = 1;}
}
</script>
<title><%=oblog.l_uname%>-用户管理后台</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="oBlogStyle/style.css" type="text/css" />
<script src="inc/main.js" type="text/javascript"></script>
<script src="oBlogStyle/menu.js" type="text/javascript"></script>
<noscript>Your browser does not support JavaScript!<br/>您的浏览器不支持JavaScript!这将使您不能使用某些功能.</noscript>
</head>
<body>
