<!--#include  file="inc/Class_UserCommand.asp"-->
<%
Dim UserId,Action,strReturn,classid
UserId=Request.QueryString("uid")
Action=LCase(Request.QueryString("do"))
if (action="index" or action="message") and Request("page")="1" then '�ж���ҳ
	Response.Write("window.location='"&action&"."&f_ext&"'")
	Response.End()
end if
select Case  Action
	Case "index","blogs","photos","month","day","message", "comment", "tag_blogs", "tag_photos", "tags", "show","album","info"
		Dim objUC
		Set objUC=New Class_UserCommand
		objUC.UserId=UserId
		strReturn=objUC.Process
		'Response.Write strReturn & VbCrlf
		Response.Write strReturn
		strReturn=objUC.CreateCalendar
		Response.Write strReturn & VbCrlf
		Set objUC=Nothing
		Set oBlog=Nothing
	Case Else
		'Response.Write "����Ĳ���"
		Response.End
End select
%>