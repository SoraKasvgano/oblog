<%
'Check Boolean

Function Check01(sValue,sDefault)

	If (sValue="0" Or sValue="1") Then
		Check01=sValue
		Exit Function
	Else
		If sDefault="" Then
			Check01=""
			Exit Function
		Else
			Check01=sDefault
		End If
	End If
End Function

Function CheckInt(sValue,sDefault)
	CheckInt=sDefault
	If Not IsNumeric(sValue) Then Exit Function
	If InStr(sValue,".") Then Exit Function
	CheckInt=Int(sValue)
End Function

Function CheckStr(sValue,iLen,sMode)
	If IsNull(sValue) Then Exit Function
	'0:¼����
	If sMode=0 Then
		'�������ַ�
		'����ֹ�ַ�
		'��鳤��
		If iLen<>"" Then CheckStr=Left(sValue,iLen)
	Else
		'�����
		CheckStr=sValue
	End If

End Function
%>