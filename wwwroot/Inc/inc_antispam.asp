<%
'inc_antiSpam
'4.0改用UBB，仅支持部分标签，所以不对连接进行处理
'仅检查回复频度
Function  antiSpam(sMode)
	'内容过滤字前面已经处理过，此处进行重复性处理
	Dim rst,rstCache,iCheck,sIP,sIP1,sql
	Set rst=Server.CreateObject("Adodb.Recordset")
	iCheck=0
	sIP=oblog.userip
	'检测IP
	sIP1=Replace(sIP,".","")
	If Not IsNumeric(sIP1) Then	sIP="0.0.0.0"
	If sIP="" Then
		antiSpam="您的IP来源被系统置疑，系统不接收您的数据!"
		Exit Function
	End If
	If sMode="1" Then
		sql = "select count(commentid) From Oblog_Comment "
		If is_sqldata=0 Then
			sql = sql & "Where datediff('n',addtime,now())<=" & oblog.CacheConfig(61)
		Else
			sql = sql & "Where addtime BETWEEN DATEADD(Hour,-1*ABS("&oblog.CacheConfig(61)&"),GETDATE()) AND GETDATE()"
		End if
		Set rst=oblog.Execute(sql & " AND addip='" & sIP & "'")
		If rst(0)>oblog.CacheConfig(62) Then
			iCheck=1
		Else
			rst.Close
			Set rst=oblog.Execute(SQL)
			If rst(0)>oblog.CacheConfig(63) Then iCheck=2
		End If
	Else
		sql = "select count(messageid) From Oblog_Message "
		If is_sqldata=0 Then
			sql = sql & "Where datediff('n',addtime,now())<=" & oblog.CacheConfig(61)
		Else
			sql = sql & "Where addtime BETWEEN DATEADD(Hour,-1*ABS("&oblog.CacheConfig(61)&"),GETDATE()) AND GETDATE()"
		End if
		Set rst=oblog.Execute(sql & " AND addip='" & sIP & "'")
		If rst(0)>oblog.CacheConfig(62) Then
			iCheck=1
		Else
			rst.Close
			Set rst=oblog.Execute(SQL)
			If rst(0)>oblog.CacheConfig(63) Then iCheck=2
		End If
	End If
	rst.Close
	Set rst=Nothing
	select Case iCheck
		Case 0
			antiSpam=""
		Case 1
			If Not oblog.ChkWhiteIP(sIP) Then
				'加入黑名单
				oblog.KillIP(sIP)
				antiSpam="因为您的一些操作对系统进行了干扰，你的IP被加入黑名单"
			Else
				antiSpam = ""
			End if
		Case 2
			antiSpam="系统暂时不允许进行回复或留言操作!"
	End select
End Function

'检查特殊标签
Function ChkCommentTag(ByVal sContent)
	Dim sBadtags,aTags,i
	sBadtags="[/url],[url, href"
	aTags=Split(sBadtags,",")
	ChkCommentTag=False
	sContent=LCase(sContent)
	For i=0 To Ubound(aTags)
		If InStr(sContent,aTags(i))>0 Then
			ChkCommentTag=True
			Exit Function
		End If
	Next
End Function
%>