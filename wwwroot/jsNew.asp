<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<!--#include file="inc/Cls_XmlDoc.asp"-->
<%

Const lockUrl = ""
'部分参考DV
'如果禁止其他站点调用，请将lockUrl赋值为你允许站点的Url
'比如lockUrl = "http://www.oblog.com.cn/"，即仅允许http://www.oblog.com.cn/的网站调用此JS
'也可以同时指定多个，以|分隔，比如lockUrl = "http://www.oblog.com.cn/|http://www.oblog.cn/"

If CheckServer(Lockurl) = False Then
	jsEcho "请勿非法调用"
	Response.End
End If

Const xmlFilePath = "xmlData/jsContent.config"

Dim mainUrl
Dim oblog
Dim xmlDoc,node
Dim action
Dim iType
Dim skin_head,skin_foot,skin_main,topN,formatTime,length,SQL
Dim show
Dim thisArr
Dim strTemp

Call Page_Load()

'页面初始化
Sub Page_Load()
	Set oblog = new class_sys
	oblog.autoupdate = False
	oblog.start

	mainUrl=Trim(oblog.CacheConfig(3))

	Set xmlDoc = New Cls_XmlDoc
	xmlDoc.Unicode = False

	If Not xmlDoc.LoadXml("xmlData/jsTemplate.config") Then
		jsEcho "模板文件不存在，无法完成操作"
		Response.End
	End If

	action = Trim (LCase(Request("action")))
	If action ="" Then
		jsEcho "参数错误"
		Response.End
	End If

	Set node = XmlDoc.NodeObj("template[@name='"&action&"']")

	If node Is Nothing Then
		jsEcho "参数错误"
		Response.End
	End If

	Dim update,updateTime
	update =  XmlDoc.AtrributeValue("template[@name='"&action&"']","update")
	updateTime =  XmlDoc.AtrributeValue("template[@name='"&action&"']","updateTime")
	iType = XmlDoc.AtrributeValue("template[@name='"&action&"']","type")
	'查询是否应该重新请求调用信息
	If checkint(update) > 0 And  IsDate(updateTime) Then
		If DateDiff("s",updateTime,Now()) > Int(update) Then
			Call updateXml()
			Call SaveXml()
		End if
	Else
		Call updateXml()
		Call SaveXml()
	End If
	'输出
	Call showJS()
End Sub
'站点统计
'类型1
Sub tongji()

	Dim rs
	Dim logtoday
	If Is_Sqldata = 0 Then
		Set rs = oblog.execute("select COUNT(logid) FROM oblog_log WHERE DATEDIFF('d',truetime,Now)=0 AND isdel=0 ")
	Else
		Set rs = oblog.execute("select COUNT(logid) FROM oblog_log WHERE truetime>=CONVERT(CHAR(10),GETDATE(),120) AND truetime < CONVERT(CHAR(10),GETDATE()+1,120) AND isdel=0 ")
	End if
	logtoday=rs(0)
	strTemp = Replace(skin_main,"$logcount$",oblog.Setup(1,0))
	strTemp = Replace(strTemp,"$commentcount$",oblog.Setup(2,0))
	strTemp = Replace(strTemp,"$messagecount$",oblog.Setup(3,0))
	strTemp = Replace(strTemp,"$usercount$",oblog.Setup(4,0))
	strTemp = Replace(strTemp,"$logtoday$",oblog.Setup(10,0))
	strTemp = Replace(strTemp,"$logyestoday$",logtoday)
	show = strTemp
	strTemp = ""
	show = skin_head &  show &  skin_foot
end Sub
'用户信息相关查询
'类型2
Sub listUser()

	SQL = node.selectSingleNode("sql").text
	length = XmlDoc.AtrributeValue("template[@name='"&action&"']","length")

	Dim i,blogname,userurl
	i=0
	thisArr = GetRows(SQL)
	For i = 0 To UBound(thisArr,2)
		If Trim(thisArr(2,i))<>"" Then
			blogname=oblog.filt_html(oblog.htm2js(thisArr(2,i),False))
		Else
			blogname=oblog.filt_html(oblog.htm2js(thisArr(0,i),False))
		End If
		If oblog.CacheConfig(5)=1 Then
			userurl="http://"&thisArr(4,i)&"."&Trim(thisArr(5,i))
		Else
			userurl=mainUrl&"go.asp?userid="&thisArr(3,i)
		End If
		strTemp = Replace(skin_main,"$userurl$",userurl)
		strTemp = Replace(strTemp,"$username$",thisArr(0,i))
		strTemp = Replace(strTemp,"$blogname$",Left(blogname,length))
		strTemp = Replace(strTemp,"$logcount$",thisArr(1,i))
		show = show & strTemp
		strTemp = ""
	Next
	show = skin_head &  show &  skin_foot
End Sub
'公告
'类型3
Sub showPlacard()
	show = Replace (skin_main,"$placard$",oblog.Setup(5,0))
	show = skin_head &  show &  skin_foot
End Sub
'分类
'类型4
Sub listClass()

	SQL = node.selectSingleNode("sql").text

	Dim i
	i=0
	thisArr = GetRows(SQL)
	For i = 0 To UBound(thisArr,2)
		strTemp = Replace(skin_main,"$classurl$",mainurl&"list.asp?classid="&thisArr(0,i))
		strTemp = Replace(strTemp,"$classname$",thisArr(1,i))
		show = show & strTemp
		strTemp = ""
	Next
	show = skin_head &  show &  skin_foot
End Sub
'日志
'类型5
Sub showLog()
	Dim i
	Dim topic, postname,posttime,userurl
	Dim isClass,isSubject
	Dim classname,subjectname
	Dim classurl,subjecturl
	Dim rstmp

	SQL = node.selectSingleNode("sql").text
	length = XmlDoc.AtrributeValue("template[@name='"&action&"']","length")
	formatTime =  XmlDoc.AtrributeValue("template[@name='"&action&"']","formatTime")
	isClass =  XmlDoc.AtrributeValue("template[@name='"&action&"']","isClass")
	isSubject =  XmlDoc.AtrributeValue("template[@name='"&action&"']","isSubject")
	i=0
	thisArr = GetRows(SQL)
	For i = 0 To UBound(thisArr,2)
		postname=Trim(thisArr(0,i))
		POSTTIME=thisArr(5,i)
		topic=oblog.filt_html(oblog.htm2js(thisArr(1,i),False))
		If oblog.CacheConfig(5) = "1" Then
			userurl="http://"&thisArr(9,i)&"."&Trim(thisArr(10,i))
		else
			userurl=mainUrl&"go.asp?userid="&thisArr(8,i)
		end if
		if Len(topic)>Int(length) Then
			topic=Left(topic,length)&"..."
		end If

		strTemp = Replace(skin_main,"$logurl$",mainurl&"go.asp?logid="&thisArr(2,i))
		strTemp = Replace(strTemp,"$topic$",OB_IIF(topic,"无题"))
		strTemp = Replace(strTemp,"$userurl$",userurl)
		strTemp = Replace(strTemp,"$postname$",OB_IIF(postname,"未知"))
		strTemp = Replace(strTemp,"$posttime$",FormatDateTime(posttime,formatTime))
		strTemp = Replace(strTemp,"$iis$",thisArr(6,i))
		strTemp = Replace(strTemp,"$commentnum$",thisArr(7,i))
		If isClass = "1" Then
			Set rstmp = oblog.execute("select id,classname from oblog_logclass where id=" & thisArr(3,i))
			If Not rstmp.EOF Then
				classname = rstmp(1)
				classurl = mainurl&"list.asp?classid="&rstmp(0)
			End If
		End If
		strTemp = Replace(strTemp,"$classurl$",classurl)
		strTemp = Replace(strTemp,"$classname$",classname)
		If isSubject = "1" Then
			Set rstmp = oblog.execute("select subjectid,subjectname from oblog_subject where subjectid=" &thisArr(4,i))
			If Not rstmp.EOF Then
				subjectname = oblog.filt_html(rstmp(1))
				subjecturl = mainurl&"blog.asp?name="&thisArr(0,i)&"&subjectid="&rstmp(0)
			End If
		End If
		strTemp = Replace(strTemp,"$subjecturl$",subjecturl)
		strTemp = Replace(strTemp,"$subjectname$",subjectname)
		show = show & strTemp
		strTemp = ""
	Next
	show = skin_head &  show &  skin_foot

End Sub
'相片
'参数6
Sub showPhoto()
	On Error Resume Next
	Dim br
	SQL = node.selectSingleNode("sql").text
	br =  XmlDoc.AtrributeValue("template[@name='"&action&"']","br")

	If br = 0 Then br = 1

	Dim rs, sReadMe,imgsrc,fso,wstr,hstr,j,preImgSrc,I
	Set fso = Server.CreateObject(oblog.CacheCompont(1))

	i=0
	thisArr = GetRows(SQL)
	For i = 0 To UBound(thisArr,2)

		If IsNull(thisArr(1,i)) Then
			sReadMe = ""
		Else
			sReadMe = oblog.filt_html(thisArr(1,i))
		End If
		imgsrc=thisArr(0,i)
		If imgsrc="" Or IsNull(imgsrc) Then
		imgsrc=proIco(thisArr(0,i),4)
		preImgSrc=imgsrc
		Else
		preImgSrc=Replace(imgsrc,right(imgsrc,3),"jpg")
		preImgSrc=Replace(preImgSrc,right(preImgSrc,len(preImgSrc)-InstrRev(preImgSrc,"/")),"pre"&right(preImgSrc,len(preImgSrc)-InstrRev(preImgSrc,"/")))
		If  Not fso.FileExists(Server.MapPath(preImgSrc)) Then
			imgsrc=proIco(thisArr(0,i),4)
			preImgSrc=imgsrc
		End If
		End If
		strTemp = Replace(skin_main,"$albumurl$",mainurl&"go.asp?albumid="&thisArr(2,i))
		strTemp = Replace(strTemp,"$imgsrc$",preImgSrc)
		strTemp = Replace(strTemp,"$readme$",sReadMe)
		If (i+1) Mod br = 0 Then
			strTemp = Replace(strTemp,"$br$","<br />")
		Else
			strTemp = Replace(strTemp,"$br$","")
		End if
		show = show & strTemp
		strTemp = ""
	Next
	show = skin_head &  show &  skin_foot
End Sub
'博客之星
'参数7
Sub showBlogStar()
	Dim iCount
	Dim i,br
	SQL = node.selectSingleNode("sql").text
	br =  XmlDoc.AtrributeValue("template[@name='"&action&"']","br")
	If br = 0 Then br = 1
	i=0
	thisArr = GetRows(SQL)

	If UBound(thisArr,2) =0 Then
		strTemp = vbcrlf & skin_head & vbcrlf
		strTemp = strTemp & skin_main
		strTemp = Replace(strTemp,"$userurl$",thisArr(0,0))
		strTemp = Replace(strTemp,"$blogurl$",mainurl&"go.asp?userid="&thisArr(4,0))
		strTemp = Replace(strTemp,"$picurl$",thisArr(1,0))
		strTemp = Replace(strTemp,"$info$",thisArr(2,0))
		strTemp = Replace(strTemp,"$blogname$",thisArr(3,0))
		strTemp = Replace(strTemp,"$tr$","")
		show = show & strTemp & skin_foot & vbcrlf
	'多图片时强制大小统一
	ElseIf UBound(thisArr,2) >0 Then
		For i = 0 To UBound(thisArr,2)
			iCount = i + 1
			strTemp = Replace(skin_main,"$userurl$",thisArr(0,i))
			strTemp = Replace(strTemp,"$blogurl$",mainurl)
			strTemp = Replace(strTemp,"$picurl$",thisArr(1,i))
			strTemp = Replace(strTemp,"$info$",thisArr(2,i))
			strTemp = Replace(strTemp,"$blogname$",thisArr(3,i))
			If iCount Mod br = 0 Then
				strTemp = Replace(strTemp,"$tr$","</tr>")
			Else
				strTemp = Replace(strTemp,"$tr$","")
			End If
			show = show & strTemp
			strTemp = ""
		Next
		If Right(show, 5) <> "</tr>" Then show = show & "	</tr>" & vbcrlf

		show = skin_head & show & skin_foot & vbcrlf
	Else
		show = "&nbsp;"
	End If

End Sub
'圈子列表
'参数8
Sub showTeam()

	Dim sIco,i,islogo

	SQL = node.selectSingleNode("sql").text

	i=0
	thisArr = GetRows(SQL)

	For i = 0 To UBound(thisArr,2)

		sIco=Proico(thisArr(2,i),1)

		strTemp = Replace(skin_main,"$ico$",sico)
		strTemp = Replace(strTemp,"$gurl$",mainurl&"group.asp?gid="&thisArr(0,i) )
		strTemp = Replace(strTemp,"$tname$",thisArr(1,i))
		strTemp = Replace(strTemp,"$count0$",thisArr(3,i))
		strTemp = Replace(strTemp,"$count1$",thisArr(4,i))
		show = show & strTemp
		strTemp = ""
	Next
	show = skin_head &  show &  skin_foot
End Sub
'圈子日志
'参数9
Sub showTeamPost()

	Dim i,istime,isuname
	SQL = node.selectSingleNode("sql").text
	length = XmlDoc.AtrributeValue("template[@name='"&action&"']","length")
	formatTime =  XmlDoc.AtrributeValue("template[@name='"&action&"']","formatTime")


	i=0
	thisArr = GetRows(SQL)

	For i = 0 To UBound(thisArr,2)

		strTemp = Replace(skin_main,"$posturl$",mainurl&"group.asp?gid="&thisArr(0,i)&"&pid="&thisArr(1,i))
		strTemp = Replace(strTemp,"$topic$",oblog.Filt_html(Left(thisArr(2,i),length)))
		strTemp = Replace(strTemp,"$addtime$",FormatDateTime(thisArr(3,i),formatTime))
		strTemp = Replace(strTemp,"$author$",thisArr(4,i))

		show = show & strTemp
		strTemp = ""
	Next
	show = skin_head &  show &  skin_foot
End Sub
'Tag
'参数10
Sub showTag()
	Dim sContent,sSql,rst,iFont,iFontSize,i,iFontFamily
	Dim order,BR,iscloud
	i=0
	SQL = node.selectSingleNode("sql").text
	br = XmlDoc.AtrributeValue("template[@name='"&action&"']","br")
	iscloud = XmlDoc.AtrributeValue("template[@name='"&action&"']","iscloud")
	thisArr = GetRows(SQL)
	For i = 0 To UBound(thisArr,2)
		If iscloud="0" Then
			strTemp = Replace(skin_main,"$tagurl$",mainurl&"tags.asp?tagid="&thisArr(0,i))
			strTemp = Replace(strTemp,"$tagname$",thisArr(1,i))
			strTemp = Replace(strTemp,"$num$",thisArr(2,i))
			show = show & strTemp
			strTemp = ""
		Else
			Dim className,FontSize,FontWeight
			iFont=thisArr(2,i)
			If iFont <= 10 Then
				className = """tag_font1"""
				FontSize = "12px"
				FontWeight = "400"
			ElseIf iFont <=30 Then
				className = """tag_font2"""
				FontSize = "14px"
				FontWeight = "400"
			ElseIf iFont <=100 Then
				className = """tag_font3"""
				FontSize = "15px"
				FontWeight = "400"
			ElseIf iFont <=200 Then
				className = """tag_font4"""
				FontSize = "16px"
				FontWeight = "600"
			ElseIf iFont <=300 Then
				className = """tag_font5"""
				FontSize = "17px"
				FontWeight = "600"
			ElseIf iFont <=450 Then
				className = """tag_font6"""
				FontSize = "18px"
				FontWeight = "600"
			ElseIf iFont <=600 Then
				className = """tag_font7"""
				FontSize = "19px"
				FontWeight = "600"
			ElseIf iFont <=800 Then
				className = """tag_font8"""
				FontSize = "20px"
				FontWeight = "600"
			ElseIf iFont <=1000 Then
				className = """tag_font9"""
				FontSize = "21px"
				FontWeight = "600"
			Else
				className = """tag_font10"""
				FontSize = "22px"
				FontWeight = "600"
			End if
			if iFontSize >800 then iFontFamily="黑体"

			strTemp = Replace(skin_main,"$tagurl$",mainurl&"tags.asp?tagid="&thisArr(0,i))
			strTemp = Replace(strTemp,"$tagname$",thisArr(1,i))
			strTemp = Replace(strTemp,"$num$",thisArr(2,i))
			strTemp = Replace(strTemp,"$className$",className)
			strTemp = Replace(strTemp,"$FontSize$",FontSize)
			strTemp = Replace(strTemp,"$FontWeight$",FontWeight)
			strTemp = Replace(strTemp,"$iFontFamily$",iFontFamily)
			show = show & strTemp
			strTemp = ""

		End If
		If (i+1) Mod br = 0 Then
			show = Replace(show,"$br$","<br />")
		Else
			show = Replace(show,"$br$","")
		End If
	Next
	show = skin_head &  show &  skin_foot
End Sub
'digg
'参数11
Sub showDigg()
	Dim sRet,Sql,rs,ClassName
	Dim arrayList,i,order
	SQL = node.selectSingleNode("sql").text
	order = XmlDoc.AtrributeValue("template[@name='"&action&"']","order")
	topN =  XmlDoc.AtrributeValue("template[@name='"&action&"']","topN")
	thisArr = GetRows(SQL)
	ReDim arrayList(topN-1)

	For i = 0 To UBound(thisArr,2)
		strTemp = Replace(skin_main,"$userurl$",mainurl&"go.asp?userid="&thisArr(5,i))
		strTemp = Replace(strTemp,"$num$",thisArr(0,i))
		If true_domain = 1 Then
			strTemp = Replace(strTemp,"$url$",thisArr(1,i))
		Else
			strTemp = Replace(strTemp,"$url$",mainurl&thisArr(1,i))
		End if
		strTemp = Replace(strTemp,"$title$",thisArr(2,i))
		strTemp = Replace(strTemp,"$addtime$",thisArr(3,i))
		strTemp = Replace(strTemp,"$username$","<span name=""nickname_"&thisArr(5,i)&""" id=""nickname_"&thisArr(5,i)&""">"&thisArr(5,i)&"</span>")
		arrayList(i) = thisArr(5,i)

		show = show & strTemp
		strTemp = ""
	Next
	show = skin_head &  show &  skin_foot
	If order = "1" Then
		show = show & oblog.GetNickNameById (arrayList,i,UBound(thisArr,2)&order)
	End if
End Sub
'被digg的用户
Sub showUserDigg()
	Dim sRet,Sql,rs
	Dim i,order
	SQL = node.selectSingleNode("sql").text
'	order = XmlDoc.AtrributeValue("template[@name='"&action&"']","order")
	topN =  XmlDoc.AtrributeValue("template[@name='"&action&"']","topN")
	thisArr = GetRows(SQL)
	For i = 0 To UBound(thisArr,2)

		strTemp = Replace(skin_main,"$userurl$",mainurl&"go.asp?userid="&thisArr(0,i))
		strTemp = Replace(strTemp,"$num$",OB_IIF(thisArr(4,i),0))
		strTemp = Replace(strTemp,"$imgsrc$",Proico(thisArr(1,i),1))
		strTemp = Replace(strTemp,"$username$",OB_IIF(thisArr(3,i),thisArr(2,i)))
		show = show & strTemp
		strTemp = ""
	Next
	show = skin_head &  show &  skin_foot
End Sub
Sub showLogin()
	Dim order
	order =  XmlDoc.AtrributeValue("template[@name='"&action&"']","order")

	strTemp = Replace(skin_main,"$blogurl$",Trim(oblog.CacheConfig(3)))
	If order = 0 Then
		strTemp = Replace(strTemp,"$n$","")
	Else
		strTemp = Replace(strTemp,"$n$","&n=1")
	End if
	show = strTemp
End Sub
'重新查询请求的调用信息
Sub updateXml()
	Call getInfo()

	Select Case itype
		Case 1 : Call tongji()
		Case 2 : Call listUser()
		Case 3 : Call showPlacard()
		Case 4 : Call listClass()
		Case 5 : Call showLog()
		Case 6 : Call showPhoto()
		Case 7 : Call showBlogStar()
		Case 8 : Call showTeam()
		Case 9 : Call showTeamPost()
		Case 10 : Call showTag()
		Case 11 : Call showDigg()
		case 12 : Call showUserDigg()
		Case 13 : Call showLogin()
		Case Else
	End Select
End Sub
'更新节点的最后更新时间并保存至XML配置文件中
Sub SaveXml()
	xmlDoc.setAttributeNode "template[@name='"&action&"']","updateTime",Now()
	xmlDoc.Save()
	Dim xmlTemp,nodeTemp
	Set xmlTemp = New Cls_XmlDoc
	xmlTemp.Unicode = False
	If Not xmlTemp.LoadXml(xmlFilePath) Then
		xmlTemp.Create "",""
		xmlTemp.InsertElement xmlTemp.NodeObj("root"),"content",show,False,True
		xmlTemp.setAttributeNode "content","name",action
	End If
	Set nodeTemp = xmlTemp.NodeObj("content[@name='"&action&"']")
	If Not ( nodeTemp Is Nothing ) Then
		xmlTemp.UpdateNodeText "content[@name='"&action&"']",show,True
	Else
		xmlTemp.InsertElement2 xmlTemp.NodeObj("root"),"content",show,True,"name",action
	End if
	xmlTemp.Save()
End Sub
'获取用户自定义的信息
Sub getInfo()
	skin_head = node.selectSingleNode("head").text
	skin_foot = node.selectSingleNode("foot").text
	skin_main = node.selectSingleNode("main").text
End Sub
'以JS方式输出文本
Sub jsEcho (ByVal str)
	Response.Write "document.write('"
	Response.Write str
	Response.Write "')"
End Sub

Sub showJs()
	If show = "" Then
		Dim xmlTemp,nodeTemp
		Set xmlTemp = New Cls_XmlDoc
		xmlTemp.Unicode = False
		If Not xmlTemp.LoadXml(xmlFilePath) Then
			Call SaveXml()
		End IF
		Set nodeTemp = xmlTemp.NodeObj("content[@name='"&action&"']")
		If nodeTemp Is Nothing Then
			Call updateXml()
			Call SaveXml()
		Else
			show = nodeTemp.Text
		End If
	End If
	If itype = 13 Then
		Response.Write show
	Else
		Response.Write oblog.htm2js(show,True)
	End if
End Sub
'查询是否为非法调用
Function CheckServer(str)
	Dim i,servername
	If str="" Then
		CheckServer = True
		Exit Function
	Else
		CheckServer = False
	End If
	str=Split(Cstr(str),"|")
	servername=Request.ServerVariables("HTTP_REFERER")
	For i=0 to Ubound(str)

		If Right(str(i),1)="/" Then str(i)=Left(Trim(str(i)),Len(str(i))-1)

		If Lcase(left(servername,Len(str(i))))=Lcase(str(i)) then
			checkserver = True
			Exit For
		Else
			checkserver = False
		End if
	Next
End Function
Function GetRows(ByVal sql)
	Dim arrTemp,rs
	Set rs=oblog.Execute(Sql)
	If Not rs.Eof Then
		arrTemp = rs.GetRows(-1)
	Else
		show = "无记录"
		Call SaveXml()
		Call showJs()
		Response.End
	End if
	Set rs = Nothing
	GetRows = arrTemp
End Function
%>