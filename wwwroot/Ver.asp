<%
Dim Ver
Dim Ver0
Dim Ver1
Dim Ver2
Dim Ver3
Ver0 = "4.60"
Ver1 = "Final"
Ver2 = "Build080827"
Ver3 = "Access"
Ver = Ver0 & " " & Ver1 & " " & Ver2 & " (" &Ver3 &")"

'------------------
	If InStr(LCase(Request.ServerVariables("SCRIPT_NAME")),"ver.asp") Then Response.Write Ver

%>
