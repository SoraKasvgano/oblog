<%
'这个是纯粹的加密解密部分.生成key的部分在tools目录里.
'Rsa加密解密类,需要事先定义RsaKeyCode值.
'用法 RsaCode("加密解密的字符串","en/de") en 是加密 de 是解密
Class clsRSA
	Public PrivateKey
	Public PublicKey
	Public Modulus
	Public Function Crypt(pLngMessage, pLngKey)
		On Error Resume Next
		Dim lLngMod
		Dim lLngResult
		Dim lLngIndex
		If pLngKey Mod 2 = 0 Then
		lLngResult = 1
		For lLngIndex = 1 To pLngKey / 2
		lLngMod = (pLngMessage ^ 2) Mod Modulus
		' Mod may error on key generation
		lLngResult = (lLngMod * lLngResult) Mod Modulus
		If Err Then Exit Function
		Next
		Else
		lLngResult = pLngMessage
		For lLngIndex = 1 To pLngKey / 2
		lLngMod = (pLngMessage ^ 2) Mod Modulus
		On Error Resume Next
		' Mod may error on key generation
		lLngResult = (lLngMod * lLngResult) Mod Modulus
		If Err Then Exit Function
		Next
		End If
		Crypt = lLngResult
	End Function

	Public Function Encode(ByVal pStrMessage)
		Dim lLngIndex
		Dim lLngMaxIndex
		Dim lBytAscii
		Dim lLngEncrypted
		lLngMaxIndex = Len(pStrMessage)
		If lLngMaxIndex = 0 Then Exit Function
		For lLngIndex = 1 To lLngMaxIndex
		lBytAscii = Asc(Mid(pStrMessage, lLngIndex, 1))
		lLngEncrypted = Crypt(lBytAscii, PublicKey)
		Encode = Encode & NumberToHex(lLngEncrypted, 4)
		Next
	End Function

	Public Function Decode(ByVal pStrMessage)
		Dim lBytAscii
		Dim lLngIndex
		Dim lLngMaxIndex
		Dim lLngEncryptedData
		Decode = ""
		lLngMaxIndex = Len(pStrMessage)
		For lLngIndex = 1 To lLngMaxIndex Step 4
		lLngEncryptedData = HexToNumber(Mid(pStrMessage, lLngIndex, 4))
		lBytAscii = Crypt(lLngEncryptedData, PrivateKey)
		Decode = Decode & Chr(lBytAscii)
		Next
	End Function

	Private Function NumberToHex(ByRef pLngNumber, ByRef pLngLength)
		NumberToHex = Right(String(pLngLength, "0") & Hex(pLngNumber), pLngLength)
	End Function

	Private Function HexToNumber(ByRef pStrHex)
		HexToNumber = CLng("&h" & pStrHex)
	End Function
End Class

Public Function Rsacode(StrMessage,t)
	Dim LngKeyE
	Dim LngKeyD
	Dim LngKeyN
	Dim ObjRSA	
	On Error Resume Next 
	LngKeyE=Split(RsaKeyCode,",",-1,1)(0)	 
	LngKeyD=Split(RsaKeyCode,",",-1,1)(1)
	LngKeyN=Split(RsaKeyCode,",",-1,1)(2)	
	If Err Then ECHO_STR "错误","取Rsa解密密钥失败！请管理员重新设置RsaKeyCode值在Config.asp里！",1	 
	If Len(StrMessage)>0 Then 
		Set ObjRSA = New clsRSA	
				If t="en" Then	
					ObjRSA.PublicKey = LngKeyE
					ObjRSA.Modulus = LngKeyN
					StrMessage=Replace(escape(StrMessage),"%u",",")
					StrMessage = ObjRSA.Encode(StrMessage)					
				ElseIf t="de" Then 
					ObjRSA.PrivateKey =LngKeyD
					ObjRSA.Modulus = LngKeyN
					StrMessage=ObjRSA.Decode(StrMessage)
					StrMessage=unescape(Replace(StrMessage,",","%u"))
				End If 
		Set ObjRSA = Nothing
	End If 
	RsaCode=StrMessage
End Function 


%>