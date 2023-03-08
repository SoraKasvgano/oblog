<!--#include file="user_top.asp"-->
<%
dim show_err,errmsg,strerr,errmsg1,i
errmsg=Trim(Request("message"))
If errmsg<>"" Then
	errmsg=Split(errmsg,"_")
	For i=0 to UBound(errmsg)
		If i=0 Then
			errmsg1=errmsg1&"<li>"&errmsg(i)
		Else
			errmsg1=errmsg1&"<br><li>"&errmsg(i)
		End If
	Next
End If
show_err= "<table border=0 align=center class=""user_prompt"">" & vbcrlf
show_err=show_err & "  <tr><td class=""user_prompt_top"">提示信息内容</td></tr>" & vbcrlf
show_err=show_err & "  <tr><td valign='top'>" & errmsg1 &"</td></tr>" & vbcrlf
show_err=show_err & "  <tr align='center'><td class=""user_prompt_end""><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
show_err=show_err & "</table>" & vbcrlf
%>
<body style="overflow:hidden;background:#fff" scroll="no">
<div id="main">
  <div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
	</div>
  </div>
  <div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
	</div>
    <div class="content_body"><%=show_err%>
	</div>
  </div>
</div>
</body>
</html>