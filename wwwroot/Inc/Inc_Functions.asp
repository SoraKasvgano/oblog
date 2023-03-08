<%
'�˴�Ϊ��Bool���жϣ�Ҳ�������ڻ�����Bool���ж�
'���Ŀ��ֵΪ�ջ���Null,��ָ��һ��Ĭ��ֵ,��ָ����Ĭ��Ϊ��
Function ob_IIF(byval var1,byval dValue)
	Dim sReturn
	If IsNull(var1) Or IsEmpty(var1) Then
		sReturn=""
	Else
		sReturn=Trim(var1)
	End If
	If sReturn="" Then sReturn=dValue
	ob_IIF=sReturn
End Function

'�˴����ڲ������жϣ����Ϊ�棬������ΪA����������ΪB
'���Ŀ��ֵΪ�ջ���ΪNull����Ĭ��Ϊfalse
Function ob_IIF2(byval var1,byval dValue1,byval dValue2)
	Dim bValue,sReturn
	If IsNull(var1) Or var1="" Then
		bValue=false
	Else
		If var1="0" or var1=false Then
			bValue=false
		Else
			bValue=true
		End If
	End If
	If bValue Then
		sReturn=dValue1
	Else
		sReturn=dValue2
	End If
	ob_IIF2=sReturn
End Function

'���ݼ�¼�����˻��ָ��ֵ
Function GetRsValue(byval rst1,field1,field2,value1,type1)
	rst1.Filter=""
	If rst1.Eof Then Exit Function
	rst1.Movefirst
	If rst1.Eof Then
		GetRsValue=""
	Else
		'��ֵ��
		If type1="0" Or type1="" Then
			rst1.Filter=field1 & "=" & value1
		'�ַ���
		Else
			rst1.Filter=field1 & "='" & value1 & "'"
		End If
		If Not rst1.Eof Then
			GetRsValue=rst1(field2)
		Else
			GetRsValue=""
		End If
	End if
End Function

'����ģʽ
Sub OB_Debug(str,iend)
Dim bugStr
	bugStr = bugStr &  "<br />---------------------------------������Ϣ��ʼ---------------------------------<br/>"
	If IsNull(str) Then
		bugStr = bugStr &  "ֵΪNull"
	ElseIf IsEmpty(str) Then
		bugStr = bugStr &  "ֵΪEmpty"
	ElseIf IsArray(str) Then
		bugStr = bugStr &  "ֵΪArray"
	Else
		If str="" Then
			bugStr = bugStr &  "ϵͳ��ʾ��ִ�е���������"
		Else
			bugStr = bugStr &  str
		End if
	End If
	bugStr = bugStr &  "<p>����ʱ��:" & Now & "</p>"
	bugStr = bugStr &  "<br/>---------------------------------������Ϣ����---------------------------------<br/>"
	ECHO_STR "Echo Debug info",bugStr,iend

End Sub

Sub ReturnClientMsg(byval divid,byval msg)
	Dim sReturn
	sReturn= "<script language=javascript>if(chkdiv("""& divid &""")==true) { document.getElementById(""" & divid &""").innerHTML="""& msg &""";}</script>"
End Sub

Function unHtml(content)
    On Error Resume Next
    unHtml = content
    If content <> "" Then
        unHtml = Server.HTMLEncode(unHtml)
        unHtml = Replace(unHtml, vbCrLf, "<br>")
        unHtml = Replace(unHtml, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
        unHtml = Replace(unHtml, " ", "&nbsp;")
        unHtml = Replace(unHtml, "&", "")
        unHtml = Replace(unHtml, "?", "")
    End If
End Function

'x<60     -Minutes
'60<=x<1440 -Hours
'x>=24 -Days
'Response.Write FmtMinutes("2006-4-30 12:21")
Function FmtMinutes(sTime)
	Dim i,j,sReturn,iMinutes
	If IsNull(sTime) Or sTime="" Then
		FmtMinutes="-"
		Exit Function
	End If
	iMinutes=Datediff("n",sTime,Now)
	If iMinutes<60 Then
		FmtMinutes=iMinutes & "����"
		Exit Function
	End If
	i=iMinutes Mod 60
	j=iMinutes \ 60
	If j<24 Then
		FmtMinutes=j & "Сʱ"' & i & "&nbsp;����"
	Else
		'Re do
		i = i Mod 24
		j = j \ 24
		FmtMinutes=j & "��"' & i & "&nbsp;Сʱ"
	End If
End Function

'------------------------------------------------
'EncodeJP(byval strContent)
'���ı���
'10k���±������С��0.01�룬����Ӱ�쵽ִ��Ч��
'Ŀǰ��Ҫ���µ�λ��Ϊ��
'վ��������ĸ�����Ŀ�����ơ�����
'��������ʱ�ı��⡢���ݡ��ؼ���
'��������/����ʱ������
'����ʱ�Թؼ��ֽ��б���
'��ʱ������ע��������
'���������������ʹ��
'------------------------------------------------
Function EncodeJP(byval strContent)

	If strContent="" Then Exit Function

	'SQL�汾�����б���
	If CBool(Is_Sqldata) Then
		EncodeJP=strContent
		Exit Function
	End If

	strContent=Replace(strContent,"��","&#12460;")
    strContent=Replace(strContent,"��","&#12462;")
    strContent=Replace(strContent,"��","&#12464;")
    strContent=Replace(strContent,"��","&#12450;")
    strContent=Replace(strContent,"��","&#12466;")
    strContent=Replace(strContent,"��","&#12468;")
    strContent=Replace(strContent,"��","&#12470;")
    strContent=Replace(strContent,"��","&#12472;")
    strContent=Replace(strContent,"��","&#12474;")
    strContent=Replace(strContent,"��","&#12476;")
    strContent=Replace(strContent,"��","&#12478;")
    strContent=Replace(strContent,"��","&#12480;")
    strContent=Replace(strContent,"��","&#12482;")
    strContent=Replace(strContent,"��","&#12485;")
    strContent=Replace(strContent,"��","&#12487;")
    strContent=Replace(strContent,"��","&#12489;")
    strContent=Replace(strContent,"��","&#12496;")
    strContent=Replace(strContent,"��","&#12497;")
    strContent=Replace(strContent,"��","&#12499;")
    strContent=Replace(strContent,"��","&#12500;")
    strContent=Replace(strContent,"��","&#12502;")
    strContent=Replace(strContent,"��","&#12502;")
    strContent=Replace(strContent,"��","&#12503;")
    strContent=Replace(strContent,"��","&#12505;")
    strContent=Replace(strContent,"��","&#12506;")
    strContent=Replace(strContent,"��","&#12508;")
    strContent=Replace(strContent,"��","&#12509;")
    strContent=Replace(strContent,"��","&#12532;")

    EncodeJP=strContent
End Function

'------------------------------------------------
'FilterJS(strHTML)
'���˽ű�
'------------------------------------------------
Function FilterJS(byval strHTML)
	Dim objReg,strContent
	If IsNull(strHTML) OR strHTML="" Then Exit Function
	Set objReg=New RegExp
	objReg.IgnoreCase =True
	objReg.Global=True
	objReg.Pattern="(&#)"
	strContent=objReg.Replace(strHTML,"")
	objReg.Pattern="(function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie)"
	strContent=objReg.Replace(strContent,"")
	objReg.Pattern="(on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
	strContent=objReg.Replace(strContent,"")
	FilterJS=strContent
	strContent=""
	Set objReg=Nothing
End Function

'------------------------------------------------
'CheckInt(byval strNumber)
'��鲢ת������ֵ
'------------------------------------------------
Function CheckInt(byval strNumber)
	If isNull(strNumber) OR Not IsNumeric(strNumber) Then
		strNumber=0
	End If
	CheckInt=CLng(strNumber)
End Function

'------------------------------------------------
'ProtectSql(sSql)
'���ڽ��յ�ַ����������ʱSQL��ϱ���
'------------------------------------------------
'��ֹSQLע��
Function ProtectSQL(sSql)
	If ISNull(sSql) Then Exit Function
	sSql=Trim(sSql)
	If sSql="" Then Exit Function
	sSql=Replace(sSql,Chr(0),"")
	sSql=Replace(sSql,"'","��")
	sSql=Replace(sSql," ","")
	sSql=Replace(sSql,"%","��")
	sSql=Replace(sSql,"-","��")
	ProtectSQL=sSql
End Function

'�����û������ĸ�����Ϣ���ˣ����໰����
Function HTMLEncode(fString)
	If Not IsNull(fString) Then
		fString = Replace(fString, ">", "&gt;")
		fString = Replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32), " ")		'&nbsp;
		fString = Replace(fString, CHR(9), " ")			'&nbsp;
		fString = Replace(fString, CHR(34), "&quot;")
		'fString = Replace(fString, CHR(39), "&#39;")	'�����Ź���
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
		'fString=ChkBadWords(fString)
		HTMLEncode = fString
	End If
End Function

'------------------------------------------------
'RemoveHtml(byval strContent)
'�Ƴ�HTML���
'��Ҫ�û����浽���ݿ�ǰ�Ĺ���
'------------------------------------------------
Function RemoveHtml(byval strContent)
	Dim objReg ,strTmp
	If strContent="" OR ISNull(strContent) Then Exit Function
	Set objReg=new RegExp
	objReg.IgnoreCase =True
	objReg.Global=True
	objReg.Pattern="<(.[^>]*)>"
	strTmp=objReg.replace(strContent, "")
	Set objReg=Nothing
	RemoveHtml=strTmp
	strTmp=""
End Function
'------------------------------------------------
'RemoveUBB(byval strContent)
'�Ƴ�UBB���
'��Ҫ�û����浽���ݿ�ǰ�Ĺ���
'------------------------------------------------
Function RemoveUBB(byval strContent)
	Dim objReg ,strTmp
	If strContent="" OR ISNull(strContent) Then Exit Function
	Set objReg=new RegExp
	objReg.IgnoreCase =True
	objReg.Global=True
	objReg.Pattern="[.+?]"
	strTmp=objReg.replace(strContent, "")
	Set objReg=Nothing
	RemoveUBB=strTmp
	strTmp=""
End Function
'------------------------------------------------
'RedirectBy301(strURL)
'��������������301�ض�����������Ŀ���ַ
'------------------------------------------------
Sub RedirectBy301(ByVal strURL)
	Response.Clear
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location",strURL
	Response.End
End Sub

'��ȡ������IP
'Response.Write GetIP
Function GetIP()
	Dim sIP
	If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
		sIP = Request.ServerVariables("REMOTE_ADDR")
	ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
		sIP = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)
	ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
		sIP = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
	Else
		sIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	End If
	GetIP = CheckIP(sIP)
	If sIP = "" Then sIP = "0.0.0.0"
End Function

Function CheckIP(sIP)
	sIP=Trim(sIP)
	sIP=Replace(sIP,".",",")
	sIP=ChkIDs(sIP)
	If sIP<>"" Then sIP=Replace(sIP,",",".")
	CheckIP=sIP
End Function

Function ChkIDs(byval sIDs)
	Dim aIDs,i,sReturn
	sIDs=Trim(sIDs)
	If Len(sIDs)=0  Then Exit Function
	aIDs=Split(sIDs,",")
	For i=0 To Ubound(aIDs)
		'�������ⲻ���ϵ��ַ���ֱ������
		If Not IsNumeric(aIDs(i)) Then
			Exit Function
		Else
			sReturn=sReturn & "," & CLng (aIDs(i))
		End If
	Next
	If Left(sReturn,1)="," Then sReturn=Right(sReturn,Len(sReturn)-1)
	ChkIDs=sReturn
	sReturn=""
End Function

Function FilterIDs(byval strIDs)
	Dim arrIDs,i,strReturn
	strIDs=Trim(strIDs)
	If Len(strIDs)=0  Then Exit Function
	arrIDs=Split(strIDs,",")
	For i=0 To Ubound(arrIds)
		If IsNumeric(arrIDs(i)) Then
			strReturn=strReturn & "," & CLng (arrIDs(i))
		End If
	Next
	If Left(strReturn,1)="," Then strReturn=Right(strReturn,Len(strReturn)-1)
	FilterIDs=strReturn
End Function

Function FilterStrings(byval strIDs)
	Dim arrIDs,i,strReturn
	strIDs=Trim(strIDs)
	If Len(strIDs)=0  Then Exit Function
	arrIDs=Split(strIDs,",")
	For i=0 To Ubound(arrIds)
		If arrIDs(i)<>"" Then
			strReturn=strReturn & "," & arrIDs(i)
		End If
	Next
	If Left(strReturn,1)="," Then strReturn=Right(strReturn,Len(strReturn)-1)
	FilterStrings=strReturn
End Function

Function RndPassword(myLength)
	Const minLength = 6
	Const maxLength = 12
	Randomize
	Dim X, Y, strPW

	If myLength = 0 Then
		Randomize
		myLength = Int((maxLength * Rnd) + minLength)
	End If


	For X = 1 To myLength
		Y = Int((3 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase

		select Case Y
			Case 1
				'Numeric character
				Randomize
				strPW = strPW & CHR(Int((9 * Rnd) + 48))
			Case 2
				'Uppercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 65))
			Case 3
				'Lowercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 97))
		End select
	Next
	RndPassword = strPW '& Int(rnd*timer)

End Function

'��ʱ��������ִ�
'0:��;1:ʱ;2:��;3:��
Function GetDateCode(sDate,sMode)
	Dim sReturn
	If Not IsDate(sDate) Or IsNull(sDate) Then sDate = Now()
	sReturn=Year(sDate) & Right("0" & Month(sDate),2) & Right("0" & Day(sDate),2)
	select Case sMode
		Case "1"
			sReturn=sReturn & Right("0" & Hour(sDate),2)
		Case "2"
			sReturn=sReturn & Right("0" & Hour(sDate),2) & Right("0" & Minute(sDate),2)
		Case "3"
			sReturn=sReturn & Right("0" & Hour(sDate),2) & Right("0" & Minute(sDate),2) & Right("0" & Second(sDate),2)
	End select
	GetDateCode=sReturn
End Function

'���ִ��ֽ�Ϊʱ��
Function DeDateCode(sDateCode)
	If IsDate(sReturn) Then DeDateCode=sDateCode:Exit Function
	Dim iLen,sReturn
	iLen=Len(sDateCode)
	select Case iLen
		Case 6
			sReturn=Left(sDateCode,4) & "-" & Right(sDateCode,2)
		Case 8
			sReturn=Left(sDateCode,4) & "-" & Mid(sDateCode,5,2) & "-" & Right(sDateCode,2)
		Case 10
			sReturn=Left(sDateCode,4) & "-" & Mid(sDateCode,5,2) & "-" & Mid(sDateCode,7,2) & " " & Right(sDateCode,2)& ":00:00"
		Case 12
			sReturn=Left(sDateCode,4) & "-" & Mid(sDateCode,5,2) & "-" & Mid(sDateCode,7,2) & " " & Mid(sDateCode,9,2) & ":" &  Right(sDateCode,2)& ":00"
		Case 14
			sReturn=Left(sDateCode,4) & "-" & Mid(sDateCode,5,2) & "-" & Mid(sDateCode,7,2) & " " & Mid(sDateCode,9,2) & ":" & Mid(sDateCode,11,2) & ":" & Right(sDateCode,2)
	End select
	DeDateCode=sReturn
End Function

Sub SystemState()
Dim CloseMsg
	If Application(cache_name_user&"_systemstate")="stop"  Then
		If Session("adminname")="" Then
			If Right(LCase(Request.ServerVariables("SCRIPT_NAME")),16)<>"/admin_login.asp" And    Not CBool(InStr(LCase(Request.ServerVariables("SCRIPT_NAME")),LCase(IncCodePath))) Then

		    If Application(cache_name_user&"_systemnote")<>"" Then
		    	CloseMsg = Application(cache_name_user&"_systemnote")
			Else
				CloseMsg = "���Ժ���ʣ�лл��"
			End If
			ECHO_STR "ϵͳ����ĳЩԭ����ʱ�ر�",CloseMsg,1

			End If
		End If
	End If
End Sub

Function GetGUID()
    Dim sRet,obj
	Set obj=Server.CreateObject("Scriptlet.Typelib")
    sRet= Mid(LCase(Replace(obj.Guid,"-","")),2,32)
    'Response.Write i &":" & sReturn & "<br>"
    Set obj=Nothing
    GetGUID=sRet
End Function

Function PageBar(total,perpage,current,filename,seed,bShow)
	'startPage:ѭ����ʼ/endPage:ѭ������/totalPage:��ҳ��
	'����URL�еĿո�
	Dim sRet,i
	sRet=""
	filename=Replace(filename," ","%20")
	Dim startPage,endPage,totalPage
	sRet= "<form name=jumpPage mothod=post action=>"
	sRet= sRet &  "<font class=tcat2>��"&total&"�� "&"ÿҳ"&perpage&"�� "

	If total mod perPage=0 Then
		totalPage=total/perPage
	Else
		totalPage=Int(total/perpage)+1
	ENd If

	If totalPage<=10 Then
		startPage=1
	Else
		If current-seed >0 Then
			startPage=current-seed
		Else
			startPage=1
		End If
	End If
	If totalPage<=10 Then
		endPage=totalPage
	Else
		If (current+seed)<totalPage Then
			endPage=current+seed
		Else
			endPage=totalPage
		End If
	End If
	if current<seed then
		if totalPage>10 THen
			endPage=10
		End If
	End if


	sRet= sRet &  "��"&current&"ҳ/��" & totalPage&"ҳ, <a href="& filename&"1>��һҳ</a> "
	if current=1 and CLng(current)<>CLng(totalPage)then
		sRet= sRet & " ��һҳ <a href="& filename&""&current+1&">��һҳ</a>"
	elseif CLng(current)>1  then
		'Response.Write Typename(current)
		If  CLng(current)< CLng(totalPage) Then
			sRet= sRet & " <a href="& filename&""&current-1&">��һҳ</a> <a href="& filename&""&current+1&">��һҳ</a>"
		elseif CLng(current)=CLng(totalPage) then
			sRet= sRet & " <a href="& filename&""&current-1&">��һҳ</a> ��һҳ"
		end if
	else
		sRet= sRet & " ��һҳ ��һҳ"
	End If
	sRet= sRet & "  <a href="& filename&totalPage&">��ĩҳ</a>"
	sRet= sRet &  "<input type=hidden name=wheretogo value=go>&nbsp;"
	'Response.write  "<input type=hidden name=wherefile value="&filename&">"
	sRet= sRet &  "  ��ת��<input name=currentPage class=border1px size=5>ҳ <input type=button value=GO class=border1px onclick='jump()'>&nbsp;"
	'Response.write  " <BR>"
	If bShow Then
		For i=startPage to endPage
			if i=cint(current) then
				sRet= sRet & "<b>"&current&"</b> "
			Else
				sRet= sRet & "<a href="&filename&i&">"&i&"</a> "
			End If
		Next
	End If
	sRet= sRet & "</font>"
	sRet= sRet & "</form>"

	sRet= sRet & "<script language=javascript>"&chr(13)
	sRet= sRet & "function jump(){"&chr(13)
	sRet= sRet & "window.location.href='"& filename & "'+document.jumpPage.currentPage.value;"&chr(13)
	sRet= sRet & "}"&chr(13)
	sRet= sRet & "</script>"&chr(13)
	PageBar=sRet
	sRet=""
End Function


function PageBarNum(total,perpage,current,filename)
	dim sRet,pageListCount,i,className
	pageListCount=10
	If total mod perPage=0 Then
		total=total/perPage
	Else
		total=Int(total/perpage)+1
	ENd If
	'Response.Write(total)
	'Response.End()
	if total>0 then
		dim startNum
		startNum=Int((current-1)/pageListCount)*pageListCount+1
		'��ʽ��Int((n-1)/col)*col+1	n�����Ĳ���	colÿ����ʾ��������		��1��ʼ��˳����
		if current<>1 then
			sRet="<span class='inactivePage'><a href='"&filename&"1' alt='��һҳ'>|&lt;</a></span>"
		end if

			if startNum-pageListCount>0 then
				sRet=sRet&"<span class='inactivePage'><a href='"&filename&""&(startNum-pageListCount)&" alt='ǰ"&pageListCount&"ҳ'>&lt;&lt</a></span>"
			end if

			for i=startNum to startNum+pageListCount-1

				if i=current then
					className="activePage"
				else
					className="inactivePage"
				end if

				sRet=sRet&"<span class='"&className&"'><a href='"&filename&i&"'>"&i&"</a></span>"

				if i>=total then
					exit for
				end if
			Next

			if startNum+pageListCount<=total then
				sRet=sRet&"<span class='inactivePage'><a href='"&filename&(startNum+pageListCount)&"' alt='��"&pageListCount&"ҳ'>&gt;&gt</a></span>"
			end if

			if current<>total then
				sRet=sRet&"<span class='inactivePage'><a href='"&filename&total&"' alt='���һҳ'>&gt;|</a></span>"
			end if
		END IF
	PageBarNum=sRet
end function

Function MakeMiniPageBar(iAll,iPer,iThis,sFileName)
	Dim sRet,i,iPages,sSeleted
	sRet=""
	sFileName=Replace(sFileName," ","%20")
	sRet= "<form name=jumpPage mothod=post action=>"
	sRet= sRet &  "��"&iAll&"��,ת�� "
	If iThis="" Or iThis="0" Then iThis=1
	If iAll mod iPer=0 Then
		iPages=iAll/iPer
	Else
		iPages=Int(iAll/iPer)+1
	End If

	sRet= sRet & "<select name=""currentPage"" onchange=""jump()"">"
	For i=1 to iPages
		If i=iThis Then
			sSeleted=" Selected"
		Else
			sSeleted=" "
		End If
		sRet= sRet & "<option value=""" & i & """" & sSeleted & ">" & i & "/" & iPages & "</option>"
	Next
	sRet= sRet & "</select></form>"
	sRet= sRet & "<script language=javascript>"&chr(13)
	sRet= sRet & "function jump(){"&chr(13)
	sRet= sRet & "window.location.href='"& sFileName & "'+document.jumpPage.currentPage.value;"&chr(13)
	sRet= sRet & "}"&chr(13)
	sRet= sRet & "</script>"&chr(13)
	MakeMiniPageBar=sRet
	sRet=""
End Function

'�����û���Ⱥ��ͷ��(sType,1-�û�,2-Ⱥ��,3-ģ��,4-���)
Function ProIco(byval sIco,byval sType)
	If IsNull(sIco) Or IsEmpty(sIco) Then sIco=""
	sIco=Trim(sIco)
	sIco=HTMLEncode(sIco)
	If sIco="" Then
		If sType="1" Then
			sIco="images/ico_default.gif"
		ElseIf sType="2" Then
			sIco="images/default_groupico.gif"
		ElseIf sType="3" Then
			sIco="images/nopic.gIf"
		ElseIf sType = "4" Then
			sIco="images/photo_default.gif"
		End If
	End If
	If Left(LCase(sico),7)<>"http://" And Left(LCase(sico),1)<>"/"  Then sico=blogurl & sico
	ProIco=sico
End Function

'������ʽ������ʽ�����뵽<head></head>��
'��ϵͳ�Զ����Head������һ��{OB_STYLE}��ǩ
'����ȡ����Style��䵽�ý�
'�����û�����/ϵͳҳ������
Function OB_PickUpCss(byref sContent)
	If sContent="" Or IsNull(sContent) Then Exit Function
	Dim oRegExp,sRet,Match,Matches
	Set oRegExp = New Regexp
	oRegExp.IgnoreCase = True
	oRegExp.Global = True

	oRegExp.Pattern = "<link.+?>"
	Set Matches =oRegExp.Execute(sContent)
	For Each Match in Matches
		sRet = sRet & Match.Value & Vbcrlf
	Next
	sContent=oRegExp.replace(sContent,"")
	oRegExp.Pattern = "\<style(.[^\[]*)\/style\>"
	Set Matches =oRegExp.Execute(sContent)
	For Each Match in Matches
		sRet = sRet & Match.Value & Vbcrlf
	Next
	sContent=oRegExp.replace(sContent,"")
	'�е����ҳ���ϵ�<body��ǩ>
	'oRegExp.Pattern = "<body>"
	'sContent =oRegExp.replace(sContent,"")
	Set oRegExp=Nothing
	OB_PickUpCss=sRet
End Function

'����OB_PickUpCss���������ٴ���
'��CSS��ȡ��ŵ�ҳ������ϲ�
Function OB_RePutCss(sContent)
	Dim sCss
	sCss=OB_PickUpCss(sContent)
	OB_RePutCss=sCss & Vbcrlf & sContent
End Function

'**************************************************
'��������AnsiToUnicode
'�� �ã�ת��Ϊ Unicode ����
'�� ����str ---- Ҫת�����ַ�
'����ֵ��ת������ַ�
'**************************************************
Public Function AnsiToUnicode(ByVal str)
	Dim i, j, c, i1, i2, u, fs, f, p
	AnsiToUnicode = ""
	p = ""
	For i = 1 To Len(str)
		c = Mid(str, i, 1)
		j = AscW(c)
		If j < 0 Then
			j = j + 65536
		End If
		If j >= 32 And j <= 128 Or j = 10 Or j = 13 Then
			If p = "c" Then
				AnsiToUnicode = " " & AnsiToUnicode
				p = "e"
			End If
			AnsiToUnicode = AnsiToUnicode & c
		Else
			If p = "e" Then
				AnsiToUnicode = AnsiToUnicode & " "
				p = "c"
			End If
			AnsiToUnicode = AnsiToUnicode & ("&#" & j & ";")
		End If
	Next
End Function

'**************************************************
'��������UnicodeToAnsi
'�� �ã�ת��Ϊ Ansi ����
'�� ����str ---- Ҫת�����ַ�
'����ֵ��ת������ַ�
'**************************************************
Function UnicodeToAnsi(ByVal str)
	If IsNull(str) or str = "" Then
		UnicodeToAnsi = ""
		Exit Function
	End If
	Dim reg,strMatch,strTemp,arrMatches
	strTemp = str
	Set reg = New RegExp
	reg.IgnoreCase = True
	reg.Global =False
	reg.Pattern = "\&#(\d*);"
	Set arrMatches = reg.Execute(str)
	For Each strMatch In arrMatches
		str = Replace(str,strMatch.Value,chrW(strMatch.SubMatches(0)))
	Next
	set reg=Nothing
	UnicodeToAnsi = str
End Function
'��ȡָ������ID�ķ�����
Function GetsubName(sid, str)
	On Error Resume Next
    Dim tmp1, tmp2,a1,a2,i
	If sid = "" Or IsNull(sid) Or sid=0 Then
        getsubname = "����"
        Exit Function
	End if
	str=Replace(str,"!!??((","##))==")
	a1=Split(str,"##))==")
	For i=0 To Ubound(a1)-1
		If i Mod 2=0 Then
			If Int(sid)=Int(a1(i)) Then
				GetsubName=a1(i+1)
				Exit Function
			End If
		End If
	Next
    getsubname = "����"
End Function

Public Function ECHO_ERR(DOTYPE,MSG,IsImportant)
	Dim xmlDoc,xmlErrPath,node,errstr
	If IsImportant=1 Then ECHO_STR "<span style=""color:#ff0000;font-weight:bold;font-size:36px;"">&#xD7;</span>���������г�������",MSG,1
End Function
Public Function ECHO_STR (title,content,isclear)
Dim BackUrl
If isclear = 1 Then Response.clear:Response.ContentType = "text/html;charset=gb2312"
If Not IsNull(Request.ServerVariables("HTTP_REFERER")) And left(lcase(Request.ServerVariables("HTTP_REFERER")),7)="http://" Then
BackUrl="  (<a href="""&Request.ServerVariables("HTTP_REFERER")&""">����</a>)"
Else
BackUrl="  (<a href=""javascript:history.go(-1)"">����</a>)"
End If
BackUrl="<span style=""font-weight:none;font-size:12px;colro:green;"">"&BackUrl&"</span>"

Response.Write "<div id=""oblog_sys_err_echo""style=""font-size:12px;color:#666;display:block;padding:20px;line-height:22px;margin:50px;background:#F2F3FB;border:1px solid #BCC8E7;text-align:left;""><div style=""font-size:14px;padding:10px 0 10px 0;font-weight:bold"">"&title&""&BackUrl&"</div><br/><b>ִ������:</b>  http://"&LCase(Request.ServerVariables("HTTP_HOST"))&LCase(Request.ServerVariables("HTTP_X_REWRITE_URL"))&"<br/><br/>"& content &"<br/>-----<br/><span style=""font-size:12px;"">�汾:"&Ver&"<span></div>"
If isclear = 1 Then Response.End
End Function

%>