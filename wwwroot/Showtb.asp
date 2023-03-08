<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<%
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>引用通告</title>
<style type="text/css">
<!--
@charset "gb2312";
* {margin: 0;padding: 0;border: 0;}
body {
color: #444;
background: #fff;
text-align: left;
margin: 0;
font-family: 'Century Gothic', Arial, Helvetica, sans-serif;
font-size: 12px;
line-height: 150%;

}
.content {
padding:0px 100px 0px 100px;

}
.content_body {
border-left: 1px #694659 solid;
border-right: 1px #694659 solid;
border-bottom: 1px #694659 solid;
background: #fff;
padding:0px;
}
.content_body h1 {
font-size:14px;
padding:6px 0px 0px 20px;
color:#f30;
font-weight:400;
}
.content_top {
padding:6px 0px 0px 15px;
width:100%;
height:30px;
background: url("images/yinyongmenu.png") no-repeat top right;
color:#fff;
font-size:14px;
font-weight:bold;
border-left: 1px #694659 solid;
border-right: 1px #694659 solid;
}
#list {
background: url("images/menubgline.png");
}
#list h1 {
padding:8px 6px 8px 6px;
font-size:14px;
color:#333;
font-weight:600;
}
#list .list_edit a {
color:#099;
font-weight:bold;
text-decoration: none;
}
#list .list_edit a:hover {
color:#f90;
font-weight:bold;
text-decoration: underline;
word-spacing:15px;
}
#list ul {
padding:1px 0px 3px 15px ;
}
-->
</style>
<script src="inc/main.js" type="text/javascript"></script>
</head>
<body>
<div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			引用通告
	</div>

    <div class="content_body">
<%
dim logid
If oblog.CacheConfig(54) = "1" Then Response.write("系统永久禁止引用通告功能!"):Response.End
If Application(cache_name_user&"_systemenmod")<>"" Then
	Dim enStr
	enStr=Application(cache_name_user&"_systemenmod")
	enStr=Split(enStr,",")
	If enStr(3)="1" Then	Response.write("系统临时禁止使用引用通告功能!"):Response.End
End if
	logid=clng(Request("id"))
	typetb()
	Set oblog=Nothing
%>
	</div>

    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>

  </div>
</body>
</html>
<%

sub typetb()
	Dim i,tburl,TBcode,nTime
	Dim rs
	tburl=oblog.CacheConfig(3)&"tb.asp?id="&logid
	'将授权码保存至数据库
	Set rs = oblog.Execute ("select TBcode FROM oblog_log WHERE logid = "&logid)
	If Not rs.EOF Then
		TBcode = rs(0)
	Else
		Response.Write "日志不存在"
		Response.End
	End if
	If TBcode = "" Or IsNull(TBcode) Then
		TBcode = GetDateCode(Now(),2) & RndPassword(12)
		oblog.Execute ("UPDATE oblog_log SET TBcode = '" &TBcode& "' WHERE logid = "&logid)
	Else
		nTime = oblog.CacheConfig(64)
		If nTime < 30 Or nTime > 1440 Then nTime = 30
		If DateDiff("n", DeDateCode(Left(TBcode, 12)), Now) > nTime Then
			TBcode = GetDateCode(Now(),2) & RndPassword(12)
			oblog.Execute ("UPDATE oblog_log SET TBcode = '" &TBcode& "' WHERE logid = "&logid)
		End If
	End If
	tburl = tburl & "&TBcode="&TBcode
%>
<h1>本文引用地址：<%=tburl%> <input type="button" onclick="copyclip('<%=tburl%>');" value="复制"></h1>
<%
	set rs = oblog.Execute("select * from oblog_trackback where logid="&logid&" order by id desc")
	if rs.eof then
%>
<ul class="list_edit">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;该日志没有被引用。</ul>
<%else%>
<h1>该日志有如下引用</h1>
<div id="list">
<%
	i = 1
	while not rs.eof
	%>
	<h1>&nbsp;&nbsp;  <a name="t<%=rs("id")%>"></a>第<%=i%>条:引用时间：<%=rs("addtime")%></h1>
	<ul class="list_edit">引用标题：<%=oblog.filt_html(rs("topic"))%></ul>
	<ul class="list_edit">引用地址：<a href="<%=rs("url")%>" target="_blank"><%=rs("url")%></a></ul>
      <%
		i = i+1
		rs.movenext
	wend
	set rs=nothing
%>
</div>
<%
	end if
end sub
%>