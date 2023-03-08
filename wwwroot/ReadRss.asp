<!--#include file="inc/inc_syssite.asp"-->
<%
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
Dim url,xml
Function BytesToBstr(strBody,CodeBase)
       dim obj
       set obj=Server.CreateObject(oblog.CacheCompont(2))
       obj.Type=1
       obj.Mode=3
       obj.Open
       obj.Write strBody
       obj.Position=0
       obj.Type=2
       obj.Charset=CodeBase
       BytesToBstr=obj.ReadText
       obj.Close
       set obj=nothing
End Function
Response.contentType="application/xml"
url=Request("feedurl")
Set xml=Server.CreateObject("Microsoft.XMLHTTP")
xml.Open "GET",url,False
xml.send
if xml.status="200" then
	Response.Write BytesToBstr(xml.responseBody,"GB2312")
end if
set xml=Nothing
Set oblog = Nothing
%>