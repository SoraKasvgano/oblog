<!-- #include file="inc/inc_syssite.asp" -->
<%
Dim Path,rs,FileID,ShowDownErr,uid,file_ext
Dim SQL
Path = Trim(Request("path"))
FileID = Trim(Request("FileID"))
If FileID ="" And Path = "" Then
	Response.Write "��������"
	Response.End
End If
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
If CheckDownLoad Then
	If Path = "" Then
		set rs = Server.CreateObject("ADODB.RecordSet")
		link_database
		SQL = ("select file_path,userid,file_ext,ViewNum FROM oblog_upfile WHERE FileID = "&CLng(FileID))
		rs.open sql,conn,1,3
		If Not rs.Eof Then
			uid = rs(1)
			file_ext = rs(2)
			rs("ViewNum") = rs("ViewNum") + 1
			rs.Update
			downloadFile Server.MapPath(rs(0)),0
		Else
			Response.Status=404
			Response.Write "�ø���������!"
		End If
		rs.Close
		Set rs = Nothing
	Else
		downloadFile Server.MapPath(Path),1
	End If
Else
	'�������ΪͼƬ�Ļ�����Ȩ�޼����޷�ͨ�������һĬ��ͼƬ����ֹ<img>����޷����ã�Ӱ����ʾЧ��
	If Path = "" Then
		Response.Status=403
		Response.Write ShowDownErr
		Response.End
	Else
		downloadFile Server.MapPath(blogdir&"images/oblog_powered.gif"),1
	End if
End if

Set oblog = Nothing

Sub downloadFile(strFile,stype)
	If InStr(strFile,Oblog.CacheConfig(56)) <= 0 Then
		Exit Sub
	End IF
	strFile  = LCase(strFile)
	If InStr(strFile,"asp") > 0 Or InStr(strFile,"mdb") > 0 Or InStr(strFile,"config")> 0 Or InStr(strFile,"js")> 0 Then
		Exit Sub
	End if
	On Error Resume Next
	Server.ScriptTimeOut=9999999
	Dim S,fso,f,intFilelength,strFilename
	strFilename = strFile
	Response.Clear
	Set s = Server.CreateObject(oblog.CacheCompont(2))
	s.Open
	s.Type = 1
	Set fso = Server.CreateObject(oblog.CacheCompont(1))
	If Not fso.FileExists(strFilename) Then
		If stype = 0 Then
			Response.Status=404
			Response.Write "�ø����Ѿ���ɾ��!"
			Exit Sub
		Else
			strFilename = Server.MapPath(blogdir&"images/nopic.gif")
		End if
	End If
	Set f = fso.GetFile(strFilename)
	intFilelength = f.size
	s.LoadFromFile(strFilename)
	If Err Then
	 	Response.Write("<h1>����: </h1>" & Err.Description & "<p>")
		Response.End
	End If
	Set fso=Nothing
	Dim Data
	Data=s.Read
	s.Close
	Set s=Nothing
	Dim ContentType
	select Case LCase(Right(strFile, 4))
	Case ".asf"
		ContentType = "video/x-ms-asf"
	Case ".avi"
		ContentType = "video/avi"
	Case ".doc"
		ContentType = "application/msword"
	Case ".zip"
		ContentType = "application/zip"
	Case ".xls"
		ContentType = "application/vnd.ms-excel"
	Case ".gif"
		ContentType = "image/gif"
	Case ".jpg", "jpeg"
		ContentType = "image/jpeg"
	Case ".wav"
		ContentType = "audio/wav"
	Case ".mp3"
		ContentType = "audio/mpeg3"
	Case ".mpg", "mpeg"
		ContentType = "video/mpeg"
	Case ".rtf"
		ContentType = "application/rtf"
	Case ".htm", "html"
		ContentType = "text/html"
	Case ".txt"
		ContentType = "text/plain"
	Case Else
		ContentType = "application/octet-stream"
	End select
	If Response.IsClientConnected Then
		If Not (InStr(LCase(f.name),".gif")>0 Or InStr(LCase(f.name),".jpg")>0 Or InStr(LCase(f.name),".jpeg")>0 Or InStr(LCase(f.name),".bmp")>0 Or InStr(LCase(f.name),".png")>0 )Then
			Response.AddHeader "Content-Disposition", "attachment; filename=" & f.name
		End If
		Response.AddHeader "Content-Length", intFilelength
 		Response.CharSet = "UTF-8"
		Response.ContentType = ContentType
		Response.BinaryWrite Data
		Response.Flush
		Response.Clear()
	End If
End Sub
'��֤����Ȩ��
Function CheckDownLoad()
	On Error Resume Next
	CheckDownLoad = False
	'����������ο����ظ���
	If oblog.CacheConfig(67) = "1" Then
		If oblog.ChkPost = False Then
			ShowDownErr = "���������Ȩ��"
			Exit Function
		Else
				CheckDownLoad = True
				Exit Function
		End If
	Else
		If oblog.CheckUserLogined = False Then
			If oblog.CacheConfig(82) = "0" Then
				ShowDownErr = "�ο����������Ȩ��,������<A HREF="""&blogurl&"reg.asp"" >ע�����û�</A>���߷�����ҳ��"
				Exit Function
			End if
		Else
			'����������Ϊ�ϴ����򷵻�True
			If uid = oblog.l_uid Then
				CheckDownLoad = True
				Exit Function
			End if
			'�����ǰ�û��鲻�������ظ���
			If oblog.l_Group(35,0) = "0" Then
				ShowDownErr = "��ǰ�û������������Ȩ��"
				Exit Function
			Else
				'������ظ�����۳�����
				If oblog.CacheScores(21) >"0" Then
					'�����ǰ����С�����ظ�����۳��Ļ���
					If oblog.l_uScores < Int(oblog.CacheScores(21)) Then
						ShowDownErr = "���ֲ��㣬���������Ȩ��"
						Exit Function
					Else
						'ִ�п۷ֲ���
						oblog.GiveScore "",-1*Abs(oblog.CacheScores(21)),""
						Session ("CheckUserLogined_"&oblog.l_uName) = ""
						Oblog.CheckUserLogined()
						ShowDownErr = ""
					End If
				End If
			End if
		End If
	End If
	If Err Then
		CheckDownLoad = False
		ShowDownErr = Err.Description
		Err.Clear
	End If
	CheckDownLoad = True
End Function
%>