<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%dim oblog,xmlStr,FeedType,isSpider
set oblog=new class_sys
oblog.autoupdate=false
oblog.start
FeedType=LCase(Trim(request("type")))
isSpider = oblog.ChkSpider(0)
'-----------------------------------------------------------------------------------------
Response.contentType="application/xml"
Response.Expires=0
On Error Resume Next 
xmlStr = xmlStr &  "<?xml version=""1.0"" encoding=""GB2312""?>"& vbcrlf
xmlStr = xmlStr & "<rss version=""2.0"">"& vbcrlf
xmlStr = xmlStr & "<channel>"& vbcrlf
xmlStr = xmlStr & "<title>"&FmtStringForXML(oblog.CacheConfig(1))&"</title>"& vbcrlf
xmlStr = xmlStr & "<link>"&oblog.CacheConfig(3)&"</link>"& vbcrlf
xmlStr = xmlStr & "<language>zh-cn</language>"& vbcrlf
xmlStr = xmlStr & "<description>"&oblog.CacheConfig(5)&"</description>"& vbcrlf
xmlStr = xmlStr & "<generator>Oblog "&oblog.Ver&"</generator>"& vbcrlf
xmlStr = xmlStr & "<webMaster>"&oblog.CacheConfig(11)&"</webMaster>"& vbcrlf
	Select Case FeedType
		Case "new"
			ShowNewLogFeed()
		Case "class"
			ShowNewLogFeed()
		Case Else
			ShowClassFeedList()
	End Select 
xmlStr = xmlStr &  "</channel>"& vbcrlf
xmlStr = xmlStr &  "</rss>"& vbcrlf
response.clear()
response.write xmlStr
response.End 

'---------------------------------------------------------------------------------------------------------------------
Sub ShowClassFeedList()
	dim rs,sql,classid,logtext

	set	rs=oblog.execute("select id,classname from oblog_logclass where idtype=0")
		if rs.Eof or rs.Bof Then
		xmlStr = xmlStr &  "<item></item>"
		Else 
		
		xmlStr = xmlStr &  "<item>" & vbcrlf
		xmlStr = xmlStr &  "<title>"&FmtStringForXML(oblog.CacheConfig(1)&"--频道列表")&"</title>" & vbcrlf
		xmlStr = xmlStr &  "<link>"&oblog.CacheConfig(3)&"</link>" & vbcrlf
		xmlStr = xmlStr &  "<author>"&FmtStringForXML(oblog.CacheConfig(1))&"</author>" & vbcrlf
		xmlStr = xmlStr &  "<pubDate>"&Now()&"</pubDate>" & vbcrlf
		xmlStr = xmlStr &  "<description><![CDATA["
		xmlStr = xmlStr &  " <a href="""&oblog.CacheConfig(3)&"rssfeed.asp?type=class&amp;classid=0""><b>在全部所有频道内的150条最新频道文章</b></a><br />"
		while not rs.Eof
			xmlStr = xmlStr &"  <a href="""&oblog.CacheConfig(3)&"rssfeed.asp?type=class&amp;classid="&rs(0)&""">在频道    <strong>"&rs(1)&"</strong>     内的50条最新频道文章</a><br />  "
			rs.MoveNext
		Wend
		End If 
	set rs=Nothing 
		xmlStr = xmlStr &  "  ]]></description>"& vbcrlf
		xmlStr = xmlStr &  "</item>"& vbcrlf

End Sub 
Sub ShowNewLogFeed()
	dim rs,sql,classid,logtext,topN
	classid=CLng(Request("classid"))
	if classid>0 then
		sql=" and classid="&classid
		topN=50
	else
		sql=""
		topN=150
	end if
	set rs=oblog.execute("select top "&topN&" topic,logfile,author,addtime,logtext,ispassword,ishide from oblog_log a  Where  (IsSpecial = 0 OR IsSpecial IS NULL) And isdel=0 and passcheck=1 and isdraft=0 and  (is_log_default_hidden=0 or is_log_default_hidden is null) "&sql&" order by logid desc")
	if rs.Eof or rs.Bof then
	  xmlStr = xmlStr &  "<item></item>"
	Else 
	while not rs.Eof
			If isSpider Then
			logtext=oblog.trueurl(rs("logtext"))
			Else 
			logtext = cutStr(rs("logtext"),500)			
			End If 
		xmlStr = xmlStr &  "<item>" & vbcrlf
		xmlStr = xmlStr &  "<title><![CDATA["&rs("topic")&"]]></title>" & vbcrlf
		if true_domain=0 then
			xmlStr = xmlStr &  "<link>"&oblog.CacheConfig(3)&rs("logfile")&"</link>" & vbcrlf
		else
			xmlStr = xmlStr &  "<link>"&rs("logfile")&"</link>" & vbcrlf
		end if
		xmlStr = xmlStr &  "<author>"&rs("author")&"</author>" & vbcrlf
		xmlStr = xmlStr &  "<pubDate>"&rs("addtime")&"</pubDate>" & vbcrlf
		xmlStr = xmlStr &  "<description><![CDATA["&logtext&"]]></description>" & vbcrlf
		xmlStr = xmlStr &  "</item>"& vbcrlf
		rs.MoveNext
	Wend
	End If 
	set rs=Nothing 
End Sub 
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
Function cutStr(str,strlen)
  Dim re
  Set re=new RegExp
  re.IgnoreCase =True
  re.Global=True
  re.Pattern="<(.[^>]*)>"
  str=re.Replace(str,"")  
  set re=Nothing
  cutStr=Replace(cutStr,"&nbsp;"," ")
  Dim l,t,c,i
  l=Len(str)
  t=0
  For i=1 to l
    c=Abs(Asc(Mid(str,i,1)))
    If c>255 Then
      t=t+2
    Else
      t=t+1
    End If
    If t>=strlen Then
      cutStr=left(str,i)&"..."
      Exit For
    Else
      cutStr=str
    End If
  Next
  cutStr=Replace(cutStr,chr(10),"")
  cutStr=Replace(cutStr,chr(13),"")
End Function

%>