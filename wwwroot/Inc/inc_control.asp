<%
Sub EchoInput(ctlName,ctlLen,ctlMaxLen,ctlValue)
	If ctlMaxLen<>"" Then ctlMaxLen="maxlength=""" & ctlMaxLen & """"
%>
	<input type="text" name="<%=ctlName%>" id="<%=ctlName%>" value="<%=ctlValue%>" size="<%=ctlLen%>" <%=ctlMaxLen%> />
<%
End Sub

Sub EchoRadio(ctlName,ctlStr1,ctlStr2,ctlValue)
	If ctlStr1="" Then ctlStr1="否"
	If ctlStr2="" Then ctlStr2="是"
	'如果未指定，则默认为否
	If ctlValue="" Then  ctlValue="0"
%>
	<input type="radio" name="<%=ctlName%>" id="<%=ctlName%>" value="0" <%If ctlValue="0" Then Response.Write "checked"%> /><%=ctlStr1%>&nbsp;
	<input type="radio" name="<%=ctlName%>" id="<%=ctlName%>" value="1" <%If ctlValue="1" Then Response.Write "checked"%> /><%=ctlStr2%>&nbsp;
<%
End Sub

Function MakeValidJs(sForm,sFunction,sValidCtl)
	Dim sRet
	sRet="<script language=""javascript"">" & vbcrlf
	sRet=sRet & "function " & sFunction & "(){" & vbcrlf
	sRet=sRet & sValidCtl
	sRet=sRet & "return true;" & vbcrlf
	sRet=sRet & "}"
	sRet=sRet & "</script>"
	Response.Write vbcrlf & sRet
End Function

Function JsValid(sForm,sCtl,cType,c1,c2,sNote)
	Dim sRet
	select Case cType
	Case "1" 'text
		'限制长度
		sRet="if (document."&sForm&"."&sCtl&".value.length<" & c1 & "||document."&sForm&"."&sCtl&".value.length>" &c2 &")" & vbcrlf
		sRet= sRet & "{" & vbcrlf
		sRet= sRet & "	alert("""& sNote &"\n长度大于" & c1 & "且小于" & c2 & """);"& vbcrlf
		sRet= sRet & "	document."&sForm&"."&sCtl&".focus();"& vbcrlf
		sRet= sRet & "	return false;"& vbcrlf
		sRet= sRet & "	}"& vbcrlf
	Case Else
		sRet="if (document."&sForm&"."&sCtl&".value.length=="""")" & vbcrlf
		sRet= sRet & "{" & vbcrlf
		sRet= sRet & "	alert("""& sNote &""");"& vbcrlf
		sRet= sRet & "	document."&sForm&"."&sCtl&".focus();"& vbcrlf
		sRet= sRet & "	return false;"& vbcrlf
		sRet= sRet & "	}"& vbcrlf
	End select
	JsValid=sRet
End Function
%>