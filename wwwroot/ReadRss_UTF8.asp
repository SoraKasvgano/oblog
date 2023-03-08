<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Dim url,http
Response.contentType="application/xml"
url=Trim(Request("feedurl"))
Set http=Server.CreateObject("Microsoft.XMLHTTP")
http.Open "GET",url,False
http.send
if http.status="200" then
	Response.Write http.responseText
end if
%>