<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%dim oblog
set oblog=new class_sys
oblog.autoupdate=false
oblog.start
Response.contentType="application/xml"
Response.Expires=0
Response.Write "<?xml version=""1.0"" encoding=""GB2312""?>"
%>
<rss version="2.0">
<channel>
<title><%=FmtStringForXML(oblog.CacheConfig(1))%></title>
<link><%=oblog.CacheConfig(3)%></link>
<description><%=oblog.CacheConfig(5)%></description>
<generator><%=oblog.Ver%></generator>
<webMaster><%=oblog.CacheConfig(11)%></webMaster>
<%
dim rs,sql,classid,logtext
classid=CLng(Request("classid"))
if classid>0 then
	sql=" and classid="&classid
else
	sql=""
end if
set rs=oblog.execute("select top 100 a.topic,a.logfile,a.author,a.addtime,a.logtext,a.ispassword,a.ishide from oblog_log a ,oblog_user b Where a.userid=b.userid AND (IsSpecial = 0 OR IsSpecial IS NULL) And a.isdel=0 and a.passcheck=1 and a.isdraft=0 and  (b.is_log_default_hidden=0 or b.is_log_default_hidden is null) "&sql&" order by a.logid desc")
if rs.Eof or rs.Bof then
  Response.write "<item></item>"
end if
while not rs.Eof
	if rs("ispassword")="" or isnull(rs("ispassword")) then
		logtext=oblog.trueurl(rs("logtext"))
	else
		logtext="此日志内容已加密"
	end if
	if rs("ishide")=1 then logtext="此日志内容已隐藏"
    Response.Write "<item>" & vbcrlf
	Response.write "<title><![CDATA["&rs("topic")&"]]></title>" & vbcrlf
	if true_domain=0 then
		Response.write "<link>"&oblog.CacheConfig(3)&rs("logfile")&"</link>" & vbcrlf
	else
		Response.write "<link>"&rs("logfile")&"</link>" & vbcrlf
	end if
	Response.write "<author>"&rs("author")&"</author>" & vbcrlf
	Response.write "<pubDate>"&rs("addtime")&"</pubDate>" & vbcrlf
 	Response.write "<description><![CDATA["&logtext&"]]></description>" & vbcrlf
	Response.write "</item>"
	rs.MoveNext
wend
set rs=nothing
%>
</channel>
</rss>
<%
Function FmtStringForXML(byval sContent)
	Dim objRegExp,strOutput
	If IsNull(sContent) Then
		FmtStringForXML=""
		Exit Function
	End If
	strOutput=Trim(sContent)
	If Instr(strOutput,"<") And Instr(strOutput,">")  Then
		'剔除<>标记
		Set objRegExp = New Regexp
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "<.+?>"
		strOutput = objRegExp.replace(strOutput, "")
		Set objRegExp = Nothing	
	End If
	strOutput=Replace(strOutput," ","")
	strOutput=Replace(strOutput,"<>","")
	strOutput=Replace(strOutput,"&nbsp;"," ")
	FmtStringForXML = strOutput	
End Function
%>