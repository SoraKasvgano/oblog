<%
'inc_antiSpam
'4.0����UBB����֧�ֲ��ֱ�ǩ�����Բ������ӽ��д���
'�����ظ�Ƶ��
Function  antiSpam(sMode)
	'���ݹ�����ǰ���Ѿ���������˴������ظ��Դ���
	Dim rst,rstCache,iCheck,sIP,sIP1,sql
	Set rst=Server.CreateObject("Adodb.Recordset")
	iCheck=0
	sIP=oblog.userip
	'���IP
	sIP1=Replace(sIP,".","")
	If Not IsNumeric(sIP1) Then	sIP="0.0.0.0"
	If sIP="" Then
		antiSpam="����IP��Դ��ϵͳ���ɣ�ϵͳ��������������!"
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
				'���������
				oblog.KillIP(sIP)
				antiSpam="��Ϊ����һЩ������ϵͳ�����˸��ţ����IP�����������"
			Else
				antiSpam = ""
			End if
		Case 2
			antiSpam="ϵͳ��ʱ��������лظ������Բ���!"
	End select
End Function

'��������ǩ
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