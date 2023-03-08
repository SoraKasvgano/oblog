<%
'Reference
'http://www.sixapart.com/pronet/docs/trackback_spec
'a ping Request might look like:
'POST http://www.example.com/trackback/5
'Content-Type: application/x-www-form-urlencoded; charset=utf-8
'title=Foo+Bar&url=http://www.bar.com/&excerpt=My+Excerpt&blog_name=Foo

Class Class_TrackBack
	Public LogId
	Public ID
	Public URL
	Public Title
	Public Blog_Name
	Public Excerpt
	Public IP
	Public Agent

	Private Function SendResult(strMsg)
		Dim strXML
		strXML="<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><response><error>%e</error><message>%m</message></response>"

		If strMsg="undiscovered" Then
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		ElseIf strMsg="repetition" Then
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		Elseif strMsg="invalid parameter" Then
			strXML=Replace(strXML,"%m",strMsg)
		Elseif strMsg="none data" Then
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		Else
			strXML=Replace(strXML,"%e","0")
			strXML=Replace(strXML,"%m",strMsg)
		End If
		Response.ContentType = "text/xml"
		Response.Clear
		Response.Write strXML

	End Function


	Public Function Receive()
		Dim UserId
		logId=CheckInt(LogId)
		IP=GetIP
		URL=ProtectSQL(URL)
		Title=ProtectSQL(Title)
		Blog_Name=ProtectSQL(Blog_Name)
		Excerpt=ProtectSQL(Excerpt)
		Response.Write logId & "<BR/>"
		Response.Write IP & "<BR/>"
		Response.Write URL & "<BR/>"
		Response.Write Title & "<BR/>"
		Response.Write Blog_Name & "<BR/>"
		Response.Write Excerpt & "<BR/>"

		If LogId=0 Then
			Call SendResult("invalid parameter")
			Receive=False
			Exit Function
		End if

		If Len(URL)=0 Then
			Call SendResult("none data")
			Receive=False
			Exit Function
		End If

		If Len(URL)>255 Then
			Call SendResult("url is long")
			Receive=False:Exit Function
		End If

		If Len(Blog_Name)>255 Then
			Call SendResult("name is long")
			Receive=False
			Exit Function
		End If
		If Len(Blog_Name)=0 Then Blog_Name="Unknow"
		If Len(Excerpt)=0 Then Excerpt=""
		If Len(Excerpt)>255 Then Excerpt=Left(Excerpt,252)&"..."
		If Len(Title)>255 Then Title=Left(Title,252)&"..."
		If Len(Title)=0 Then Title=URL

		Dim rst
		Set rst=conn.Execute("select * From [oblog_log] Where LogId=" & LogId)
		If rst.Eof Then
			Call SendResult("undiscovered")
			Exit Function
		End If
		Userid=rst("userid")
		rst.Close
		'Set rst=Nothing
		'Response.Write URL & "<br/>"
		'Response.Write ("select * From [oblog_TrackBack] Where [LogId]=" & LogId & " and tb_url='" & URL & "'")
		'Response.End()
		'Ping
		rst.Open "select * From [oblog_TrackBack] Where [LogId]=" & LogId & " and tb_url='" & URL & "'",conn,1,3
		If Not rst.bof Then
			rst.close
			Call SendResult("repetition")
			Exit Function
		Else
			rst.AddNew
			'rst("userId")=UserId
			rst("logId")=LogId
			rst("topic")=Title
			rst("Blog_Name")=Blog_Name
			rst("Excerpt")=Excerpt
			rst("Url")=URL
			rst("IP")=IP
			rst.Update
		End If
		rst.Close
		Set rst=Nothing
		conn.execute("update [oblog_log] set trackbacknum=trackbacknum+1 where logid="&LogId)
		Call SendResult("succeed")

		Receive=True

	End Function

	Public Function DeleteTrackBack()
		If IsNumeric(ID) Then LogId=CLng(ID)
		conn.Execute("Delete "& delchar & " From [oblog_TrackBack] Where id =" & ID)
		DeleteTrackBack=True
	End Function

	Private Function Ping(ByVal strTarget)
		Dim strSend,objPing
		strSend = "title=" & Server.URLEncode(Title) & "&url=" & Server.URLEncode(URL) & "&excerpt=" & Server.URLEncode(Excerpt) & "&blog_name=" & Server.URLEncode(Blog_Name)

		Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP"&MsxmlVersion)
		objPing.open "POST",strTarget,False
		objPing.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objPing.send strSend
'		OB_DEBUG objPing.responsetext,1
		If objPing.readystate = 4 Then
			If objPing.status = 200 Then
				Ping = True
			Else
				Ping = objPing.status
			End If
		Else
			Ping =  objPing.readystate
		End if
		Set objPing = Nothing
	End Function

	Public Function ProcessMultiPing(strTargetUrls)
		Dim aUrls,i,strUrl,rst
		strTargetUrls=Trim(strTargetUrls)
		If strTargetUrls="" Then Exit Function
		aUrls=Split(strTargetUrls,VBCRLF)
		For i=0 To UBound(aUrls)
			strUrl=Lcase(aUrls(i))
			If Left(strUrl,7)="http://" Then
				ProcessMultiPing =  Ping(strUrl)
			End If
			If i+1> Int(oblog.CacheConfig(74)) Then Exit Function
		Next
	End Function

	Function CheckTB(TBcode0)
		'检验引用通告授权码是否过期
		CheckTB = False
'		OB_DEBUG TBcode0,1
		If Len(TBcode0) <> 24 Then Exit Function
		Dim TBcode,nTime
		Dim rs
		nTime = oblog.CacheConfig(64)
		Set rs = oblog.Execute ("select TBcode FROM oblog_log WHERE logid = "&logid)
		If Not rs.EOF Then
			TBcode = rs(0)
		Else
			Call SendResult("log not exist")
			Exit Function
		End if
		If TBcode = "" Or IsNull(TBcode) Then
			Call SendResult("invalid tbcode")
			Exit Function
		End if
		If nTime < 30 Or nTime > 1440 Then nTime = 30
		If DateDiff("n", DeDateCode(Left(TBcode, 12)), Now) > nTime Then
			TBcode = GetDateCode(Now(),2) & RndPassword(12)
			oblog.Execute ("UPDATE oblog_log SET TBcode = '" &TBcode& "' WHERE logid = "&logid)
		End If
		If TBcode0 <> LCase(TBcode) Then
			Call SendResult("no authority")
			Exit Function
		End If
		CheckTB = True
	End Function
End Class
%>