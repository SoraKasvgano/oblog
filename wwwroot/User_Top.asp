<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="API/Class_API.asp" -->
<%
'------------------------------------------------
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
Dim t,tName,nowurl,cssfile,actiontype,url

url=Request.QueryString("url")
t=Request("t")
actiontype=Request("actiontype")
nowurl=oblog.NowUrl
dim uurl

G_P_PerMax=20
'�˳�ϵͳ
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
        tName = "��־"
    Case 1
        tName = "���"
    Case 2
        tName = "ͨѶ¼"
    Case 3
        tName = "����"
    Case Else
        t = 0
        tName = "��־"
End select
if not oblog.checkuserlogined() then
	Response.Redirect("login.asp?fromurl="&Replace(oblog.GetUrl,"&","$"))
else
	if oblog.l_ulevel=6 then
		oblog.adderrstr("��δͨ������Ա��ˣ����ܽ����̨")
		oblog.showerr
	end if
end if
'�������˿��ƽ�ֹҳ�浥������
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
<title><%=oblog.l_uname%>-�û������̨</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="oBlogStyle/style.css" type="text/css" />
<script src="inc/main.js" type="text/javascript"></script>
<script src="oBlogStyle/menu.js" type="text/javascript"></script>
<noscript>Your browser does not support JavaScript!<br/>�����������֧��JavaScript!�⽫ʹ������ʹ��ĳЩ����.</noscript>
</head>
<body>
