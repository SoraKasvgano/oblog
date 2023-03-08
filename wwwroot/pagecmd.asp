<!--#include file="inc/Class_UserCommand.asp"-->
<%
Dim UserId,Action,strReturn,classid,FileID
Dim objUC
UserId=Request.QueryString("uid")
FileID=Request.QueryString("fileid")
Action=LCase(Request.QueryString("do"))
if (action="index" or action="message") and Request("page")="1" then '判断首页
	Response.Write("window.location='"&action&"."&f_ext&"'")
	Response.End()
end if
select Case  Action
	Case "index","blogs","photos","month","day","message", "comment", "tag_blogs", "tag_photos", "tags", "show","album","info","photocomment","flash"
		Set objUC=New Class_UserCommand
		objUC.UserId=UserId
		If FileID > 0 Then objUC.FileID=FileID
		strReturn=objUC.Process
		'Response.Write strReturn & VbCrlf
		Response.Write strReturn
		strReturn=objUC.CreateCalendar
		Response.Write strReturn & VbCrlf
		Set objUC=Nothing
		Set oBlog=Nothing
	Case Else
		Response.Write oblog.htm2js_div("错误的参数","oblog_usercontent")
		Response.End
End select
%>