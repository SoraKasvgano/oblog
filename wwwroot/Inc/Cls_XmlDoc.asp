<%
Class Cls_XmlDoc
	Private fErrInfo,fxmlFile,fNodeObj,IsUnicode,froot
	Public XmlDoc
	'类初始化
	Private Sub Class_Initialize()
		On Error Resume Next
		Set XmlDoc=Server.CreateObject("Msxml2.DOMDocument"&MsxmlVersion)
		XmlDoc.preserveWhiteSpace=True
	End Sub
	'类释放
	Private Sub Class_Terminate()
		On Error Resume Next
		If IsObject(fNodeObj) Then Set fNodeObj = Nothing
		If IsObject(NodeObj) Then Set NodeObj = Nothing
		Set XmlDoc=nothing
	End Sub
	Public Property Let Unicode(ByVal Values)
		IsUnicode = Values
	End Property
	Public Property Get Unicode
		Unicode = IsUnicode
		If Unicode = "" Then Unicode = True
	End Property
	'返回一个节点的OBJ
	Public Property Get NodeObj(ByVal Values)
		Values = "//"&Values
		Set NodeObj = XMLDOC.selectSingleNode(Values)
	End Property
	'获取当前操作节点的XML
	Public Property Get NodeXml
		NodeXml = fNodeObj.xml
	End Property
	'或者当前操作XML文件的XML
	Public Property Get Xml
		Xml = XmlDoc.Xml
	End Property
	'读取最后的错误信息
	Public Property Get ErrInfo
		ErrInfo = fErrInfo
	End Property
	'根节点的名称
	Public Property Let root(ByVal Values)
		root = Values
		froot = root
	End Property
	'判断XmlDoc操作是否出现错误
	Private Function IsError()
		If XmlDoc.ParseError.errorcode<>0 Then
			fErrInfo="<h1>Error"&XmlDoc.ParseError.errorcode&"</h1>"
			fErrInfo=fErrInfo&"<B>Reason :</B>"&XmlDoc.ParseError.reason&"<br>"
			fErrInfo=fErrInfo&"<B>URL &nbsp; &nbsp;:</B>"&XmlDoc.ParseError.url&"<br>"
			fErrInfo=fErrInfo&"<B>Line &nbsp; :</B>"&XmlDoc.ParseError.line&"<br>"
			fErrInfo=fErrInfo&"<B>FilePos:</B>"&XmlDoc.ParseError.filepos&"<br>"
			fErrInfo=fErrInfo&"<B>srcText:</B>"&XmlDoc.ParseError.srcText&"<br>"
			IsError=True
		Else
			IsError = False
		End If
	End Function
	'装载一个xml文档，文档名可为空
	Function LoadXml(xmlSourceFile)
		LoadXml = False
		Dim xmlFile
		If froot = "" Then froot = "root"
		If xmlSourceFile <>"" Then
			xmlFile=Server.Mappath(Trim(xmlSourceFile))
			fxmlFile = xmlFile
		End if
		XMLDOC.async = False
		If xmlFile <>"" Then
			If XMLDOC.load(xmlFile) Then
				LoadXml = True
			End If
		End if
	End Function
	'获取当前操作节点的各种属性
	Public Property Get SelectXmlNode(ByVal NodeName,ByVal sType)
		On Error Resume Next
		NodeName = "//"&NodeName
		Set fNodeObj = XMLDOC.selectSingleNode(NodeName)
		select Case sType
			Case 0
				'节点名称
				selectXmlNode = fNodeObj.nodeName
			Case 1
				'节点TEXT值
				selectXmlNode = fNodeObj.text
			Case 2
				'节点型态(字符串)
				selectXmlNode = fNodeObj.nodeTypeString
			Case 3
				'节点型态(数字)
				selectXmlNode = fNodeObj.nodeType
			Case Else

		End select
	End Property
	'获取当前操作节点的某一属性值
	Public Property Get AtrributeValue(ByVal NodeName,ByVal atrributename)
		On Error Resume Next
		NodeName = "//"&NodeName
		Set fNodeObj = XMLDOC.selectSingleNode(NodeName)
		AtrributeValue=fNodeObj.GetAttributeNode(atrributename).Nodevalue
'		AtrributeValue=fNodeObj.GetAttribute(atrributename)
	End Property
	'创建一默认XML文档
	Function Create(byVal RootNodeName,byVal XslUrl)
		Dim PINode,RootElement
		If Trim(RootNodeName)="" Then RootNodeName="root"
		Set PINode=XmlDoc.CreateProcessingInstruction("xml", "version=""1.0""  encoding=""GB2312""")
		XmlDoc.appendChild PINode
		If XslUrl <>"" Then
			Set PINode=XmlDoc.CreateProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href="""&XslUrl&"""")
			XmlDoc.appendChild PINode
		End if
		Set RootElement=XmlDoc.createElement(Trim(RootNodeName))
		XmlDoc.appendChild RootElement
		Set PINode = Nothing
		Set RootElement = Nothing
	End Function
	'保存打开过的文件，只要保证fxmlFile不为空就可以实现保存
	Function Save()
		On Error Resume Next
		Save = False
		If fxmlFile="" Then Exit Function
		XmlDoc.Save fxmlFile
		Save=(Not IsError)
		If Err.number<>0 then
			Err.clear
			Save=False
		End If
	End Function
	'保存操作完成后的XML文档到指定位置
	Function SaveAs(ByVal SavexmlSourceFile)
		On Error Resume Next
		SaveAs = False
		If SavexmlSourceFile="" Then Exit Function
		SavexmlSourceFile = Server.MapPath(SavexmlSourceFile)
		XmlDoc.Save SavexmlSourceFile
		SaveAs=(Not IsError)
		If Err.number<>0 then
			Err.clear
			SaveAs=False
		End If
	End Function
	'修改当前操作节点的TEXT值
	Function UpdateNodeText(ByVal NodeName,byVal NewElementText,byVal IsCDATA)
		Dim ElementName
		ElementName = "//"&NodeName
		If Unicode Then
			NewElementText = AnsiToUnicode (NewElementText)
		End If
		NewElementText = Replace (NewElementText,"]]>","]]&gt;")
		Set fNodeObj = XMLDOC.selectSingleNode(ElementName)
		If fNodeObj Is Nothing Then
			'如果节点不存在则创建
			InsertElement XMLDOC.selectSingleNode(froot),NodeName,NewElementText,False,IsCDATA
			Exit Function
		End if
		Dim TextSection
		If IsCDATA Then
			Set TextSection=XmlDoc.createCDATASection(NewElementText)
			If fNodeObj.firstchild Is Nothing Then
				fNodeObj.appendChild TextSection
			Else
				fNodeObj.replaceChild TextSection,fNodeObj.firstchild
			End If
		Else
			fNodeObj.Text=NewElementText
		End If
		Set TextSection = Nothing
	End Function
	'修改当前操作节点的TEXT值
	Function UpdateNodeText2(ByVal OBJ,byVal NewElementText,byVal IsCDATA)
		If Unicode Then
			NewElementText = AnsiToUnicode (NewElementText)
		End if
		Set fNodeObj = OBJ
		If fNodeObj Is Nothing Then
			'如果节点不存在则创建
			InsertElement XMLDOC.selectSingleNode(froot),NodeName,NewElementText,False,IsCDATA
			Exit Function
		End if
		Dim TextSection
		If IsCDATA Then
			Set TextSection=XmlDoc.createCDATASection(NewElementText)
			If fNodeObj.firstchild Is Nothing Then
				fNodeObj.appendChild TextSection
			Else
				fNodeObj.replaceChild TextSection,fNodeObj.firstchild
			End If
		Else
			fNodeObj.Text=NewElementText
		End If
		Set TextSection = Nothing
	End Function
	'插入在BefelementOBJ下面一个名为ElementName，Value为ElementText的子节点。
	'IsFirst：是否插在第一个位置；IsCDATA：说明节点的值是否属于CDATA类型
	Function InsertElement(byVal BefelementOBJ,byVal ElementName,byVal ElementText,byVal IsFirst,byVal IsCDATA)
		Dim Element,TextSection
		If Unicode Then
			ElementName = AnsiToUnicode(ElementName)
		End if
		Set Element=XmlDoc.CreateElement(Trim(ElementName))
		If IsFirst Then
			BefelementOBJ.InsertBefore Element,BefelementOBJ.firstchild
		Else
			BefelementOBJ.appendChild Element
		End If
		If IsCDATA Then
			set TextSection=XmlDoc.createCDATASection(ElementText)
			Element.appendChild TextSection
		ElseIf ElementText<>"" Then
			Element.Text=ElementText
		End If
		Set Element = Nothing
		Set TextSection = Nothing
	End Function
	'插入在BefelementOBJ下面一个名为ElementName，Value为ElementText的子节点。
	'IsFirst：是否插在第一个位置；IsCDATA：说明节点的值是否属于CDATA类型
	'同时给当前新增的节点设定一个属性以及给属性复制
	Function InsertElement2(byVal BefelementOBJ,byVal ElementName,byVal ElementText,byVal IsCDATA,byVal AttributeName,byVal AttributeText)
		Dim Element,TextSection
		If Unicode Then
			ElementName = AnsiToUnicode(ElementName)
		End if
		Set Element=XmlDoc.CreateElement(Trim(ElementName))

		BefelementOBJ.appendChild Element

		If IsCDATA Then
			set TextSection=XmlDoc.createCDATASection(ElementText)
			Element.appendChild TextSection
		ElseIf ElementText<>"" Then
			Element.Text=ElementText
		End If

		Dim AttributeNode
		Set AttributeNode=Element.attributes.getNamedItem(AttributeName)
		If AttributeNode Is nothing Then
			Set AttributeNode=XmlDoc.CreateAttribute(AttributeName)
			Element.setAttributeNode AttributeNode
		End If
		AttributeNode.text=AttributeText
		Set AttributeNode = Nothing

		Set Element = Nothing
		Set TextSection = Nothing
	End Function
	'在当前操作节点上插入或修改名为AttributeName，值为：AttributeText的属性
	'如果已经存在名为AttributeName的属性对象，就进行修改。
	Function setAttributeNode(ByVal NodeName,byVal AttributeName,byVal AttributeText)
		NodeName = "//"&NodeName
		Set fNodeObj = XMLDOC.selectSingleNode(NodeName)
		Dim AttributeNode
		Set AttributeNode=fNodeObj.attributes.getNamedItem(AttributeName)
		If AttributeNode Is nothing Then
			Set AttributeNode=XmlDoc.CreateAttribute(AttributeName)
			fNodeObj.setAttributeNode AttributeNode
		End If
		AttributeNode.text=AttributeText
		Set AttributeNode = Nothing
	End Function
	'删除子节点的一个属性
	Function removeAttributeNode(ByVal NodeName,byVal AttributeName)
		NodeName = "//"&NodeName
		Set fNodeObj = XMLDOC.selectSingleNode(NodeName)
		Dim AttributeOBJ
		removeAttributeNode=false
		Set AttributeOBJ=fNodeObj.attributes.getNamedItem(AttributeName)
		If Not AttributeOBJ Is nothing Then
			fNodeObj.removeAttributeNode(AttributeOBJ)
			removeAttributeNode=True
		End If
		Set AttributeOBJ = Nothing
	End Function
	'删除一个子节点
	Function removeChild(ByVal NodeName)
		NodeName = "//"&NodeName
		Set fNodeObj = XMLDOC.selectSingleNode(NodeName)
		removeChild=False
		If Lcase(fNodeObj.nodeTypeString)="element" Then
			If fNodeObj.parentNode Is Nothing Then
				XmlDoc.removeChild(fNodeObj)
				removeChild=True
			Else
				fNodeObj.parentNode.removeChild(fNodeObj)
				removeChild=True
			End If
		End If
	End Function
End Class
%>