<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%
dim oblog
set oblog=new class_sys
oblog.start
if not oblog.checkuserlogined() then
	Response.Redirect("login.asp")
end if

if Request("action")="saversslist" then
	call saversslist()
else
	call savelog()
end if

sub savelog()
	dim tablename,sql,filetype
	dim rs,strLine
	dim sdate,edate,nurl
	nUrl=Trim("http://" & Request.ServerVariables("SERVER_NAME"))
	nUrl=lcase(nUrl & Request.ServerVariables("SCRIPT_NAME"))
	nurl=left(nUrl,instrrev(nUrl,"/"))
	filetype = lcase(Trim(Request("filetype")))
	sdate=oblog.filt_badstr(Request("selecty1")&"-"&Request("selectm1")&"-"&Request("selectd1"))
	edate=oblog.filt_badstr(Request("selecty2")&"-"&Request("selectm2")&"-"&Request("selectd2"))
	if not isdate(sdate) then oblog.adderrstr("开始日期不正确")
	if not isdate(edate) then oblog.adderrstr("结束日期不正确")
	if oblog.errstr<>"" then
		oblog.showusererr
	end if
	tablename = sdate&"到"&edate&"的日志"
	if is_sqldata=1 then
		sql="select topic,addtime,logtext from oblog_log where userid="&oblog.l_uid&" and addtime<='"&dateadd("d",1,edate)&"' and addtime>='"&sdate&"'"
	else
		sql="select topic,addtime,logtext from oblog_log where userid="&oblog.l_uid&" and addtime<=#"&dateadd("d",1,edate)&"# and addtime>=#"&sdate&"#"
	end if
	Set rs = oblog.Execute(sql)
	if filetype="xml" then
		Response.contenttype="text/xml"
		Response.Charset = "gb2312"
		Response.AddHeader "Content-Disposition", "attachment;filename="&tablename&".xml"
		Response.write "<?xml version=""1.0"" encoding=""gb2312""?>"
		Response.write vbnewline&"<rss version=""2.0"">"&vbnewline&"<channel>"
		strLine=""
		While not rs.EOF
			strLine= vbnewline&chr(9)&"<item>"
			strLine=  strLine &"<title>"&rs(0)&"</title>"
			strLine=  strLine & "<PubDate>"&rs(1)&"</PubDate>"
			strLine=  strLine &"<description><![CDATA["& newurl(rs(2),nurl) &"]]></description>"		
			strLine=  strLine &"</item>"	
			Response.write strLine
			rs.MoveNext
		Wend
		Response.write vbnewline&"</channel>"&vbnewline&"</rss>"
	elseif filetype="txt" then
		Response.contenttype="text"
		Response.AddHeader "Content-Disposition", "attachment;filename="&tablename&".txt"
		While not rs.EOF
			strLine=""
			strLine=strLine & "日志标题："&rs(0) & vbnewline
			strLine=strLine & "发表时间："&rs(1) & vbnewline
			strLine=strLine & "日志内容："&Trim(stripHTML(rs(2)))
			Response.write strLine & vbnewline & vbnewline
			rs.MoveNext	
		Wend
	else
	if filetype="htm" then
			Response.contenttype="application/ms-download"
			Response.AddHeader "Content-Disposition", "attachment;filename="&tablename&".htm"
	end if
		
		While not rs.EOF
		strLine= ""
		Response.write chr(9)&"<table style='font-size:10pt;TABLE-LAYOUT: fixed; WORD-BREAK: break-all' width='98%'align='center' bgColor=#ffffff border=1 >"& vbnewline
		Response.write chr(9)&"<tr>"& vbnewline
		strLine= strLine&chr(9)&chr(9)&"<td>"
		strLine= strLine&"日志标题："&rs(0)&"<br>"& vbnewline
		strLine= strLine&"发表时间："&rs(1)&"<br>"& vbnewline
		strLine= strLine&newurl(rs(2),nurl) &"</td>"& vbnewline
		Response.write strLine
		Response.write chr(9)&"</tr>"& vbnewline
		Response.write "</table><br>"& vbnewline
		rs.MoveNext
		Wend
	end if
	Set rs=nothing
end sub

sub saversslist()
	dim rsSubject,rs,m,ostr
	Response.contenttype="text/xml"
	Response.Charset = "gb2312"
	Response.AddHeader "Content-Disposition", "attachment;filename=rsslist.opml"
	Response.write "<?xml version=""1.0"" encoding=""gb2312""?>"
	Response.write vbnewline&"<opml version=""1.0"">"&vbnewline&"<body>"&vbnewline
	Set rsSubject = oblog.Execute("select subjectid,subjectname from oblog_subject where userid=" & oblog.l_uId & " And subjecttype=3 order by ordernum")
	set rs=oblog.execute("select * from oblog_myurl where subjectid>0 and userid="&oblog.l_uid&" order by subjectid desc")
	m=0
	while not rsSubject.eof
		if m=1 then ostr="</outline>" else ostr=""
		Response.Write ostr&"<outline title="""&rsSubject("subjectname")&""" expanded=""1"" text="""&rsSubject("subjectname")&""">"&vbnewline
		m=1
		while not rs.eof
			if rs("subjectid")=rsSubject("subjectid") then
				Response.Write "<outline xmlUrl="""&fullrssurl(rs("url"))&""" title="""&rs("title")&""" expanded=""0"" text="""&rs("title")&""" />"&vbnewline
			end if
			rs.movenext
		wend
		if not rs.eof then	rs.movefirst
		rsSubject.movenext
	wend 
	set rs=oblog.execute("select * from oblog_myurl where subjectid=0 and userid="&oblog.l_uid)
	if  not rs.eof then
		Response.Write "<outline title=""未分类"" expanded=""1"" text=""未分类"">"&vbnewline
		while not rs.eof
			Response.Write "<outline xmlUrl="""&fullrssurl(rs("url"))&""" title="""&rs("title")&""" expanded=""0"" text="""&rs("title")&""" />"&vbnewline
			rs.movenext
		wend
		Response.Write("</outline>"&vbnewline)
	end if
	Response.Write("</body></opml>")
	set rs=nothing
	set rsSubject=nothing
end sub

Function stripHTML(strHTML)
  Dim objRegExp, strOutput
  Set objRegExp = New Regexp
  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<.+?>"
  strOutput = objRegExp.replace(strHTML, "")
  strOutput = Replace(strOutput, "<", "<")
  strOutput = Replace(strOutput, ">", ">")
  stripHTML = Replace(strOutput,"&nbsp;","")
  Set objRegExp = Nothing
End Function

Function newurl(strContent,byval url)
    dim tempReg
    set tempReg=new RegExp
    tempReg.IgnoreCase=true
    tempReg.Global=true
    tempReg.Pattern="(^.*\/).*$"'含文件名的标准路径
    Url=tempReg.replace(url,"$1")
    tempReg.Pattern="((?:src|href).*?=[\'\u0022](?!ftp|http|https|mailto))"
    newurl=tempReg.replace(strContent,"$1"+Url)
    set tempReg=nothing
end Function

function fullrssurl(url)
	dim nurl
	nurl=Trim("http://" & Request.ServerVariables("SERVER_NAME"))
	if left(url,7)<>"http://" then
		fullrssurl=nurl&url
	else
		fullrssurl=url
	end if
end function
%>