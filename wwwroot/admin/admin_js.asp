<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/Cls_xmlDoc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>JS调用管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">JS调用管理：</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr >
      <td width="70" height="30"><strong>注意事项：</strong></td>
      <td height="30" style="color:red;">①添加调用后，在列表中点击相应的预览即可看到效果，将调用代码复制到任意位置即可调用。<br>②建议将时间间隔设置的稍微长一点，有利减少资源消耗</td>
    </tr>
    <tr >
      <td width="70" height="30"><strong>管理导航：</strong></td>
      <td height="30"><a href="?action=add">添加JS调用</a> | <a href="?">JS调用列表</a></td>
    </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%
Dim getType

Dim xmlDoc

ReDim getType(13)

getType(0) = ""
getType(1) = "站点统计"
getType(2) = "用户信息"
getType(3) = "站点公告"
getType(4) = "系统分类"
getType(5) = "日志"
getType(6) = "相片"
getType(7) = "博客之星"
getType(8) = "圈子列表"
getType(9) = "圈子日志"
getType(10) = "标签（TAG）"
getType(11) = "用户推荐（DIGG）日志"
getType(12) = "被推荐（DIGG）用户信息"
getType(13) = "登录窗口"

Dim eName,Intro,eType,Update,FormatTime
Dim isModify
Dim action
Dim node
Dim head,skinmain,foot
Dim Sql

eName = Trim(Request("eName"))
Intro = Trim(Request("Intro"))
eType = Trim(Request("eType"))
Update = Trim(Request("Update"))
FormatTime = Trim(Request("FormatTime"))
isModify = Trim(Request("modify"))

head = Trim(Request("head"))
skinmain = Trim(Request("main"))
foot = Trim(Request("foot"))

Dim topN,length,order,isbest

topN = Trim(request("topn"))
If topN<>"" Then topN = CLng(topN) Else topN=10
If topN  >50 Then topN = 50
length = Trim(request("length"))
If length<>"" Then length = CLng(length) Else length=20
order = Trim(request("order"))
isbest = Trim(request("isbest"))
action = Trim(Request("action"))

Select Case Trim(Request("action"))
	Case "del":Call delNode()
	Case "add","modify":Call add()
	Case "saveadd":Call saveadd()
	Case Else :Call main()
End Select

'部分参考DV
Sub main()%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">JS调用列表</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form action="" method="post" name="myform">
<table cellpadding="0" cellspacing="1" border="0" align="center" width="100%" class="border">
	<tr class="title">
		<td align="center" width="28" height="23"><strong>选中</strong></td>
		<td align="center" width="100"><strong>分类</strong></td>
		<td align="center"><strong>名称</strong></td>
		<td align="center" width="150"><strong>说明</strong></td>
		<td align="center" width="80"><strong>添加更新时间</strong></td>
		<td align="center" width="60"><strong>添加者</strong></td>
		<td align="center" width="60"><strong>操作</strong></td>
	</tr>
<%
Dim xmlDoc
Set XmlDoc=CreateObject("Msxml2.DOMDocument"&MsxmlVersion)
	If Not xmlDoc.Load(Server.Mappath("../xmlData/jsTemplate.config")) Then
		Response.Write  "模板文件不存在，无法完成操作"
		Response.End
	End If
Dim nodes,i,node
Set nodes = xmlDoc.getElementsByTagName("template")
For Each node In nodes
%>
	<tr>
		<td align=center><input name='ename' type='checkbox' id="ename" value='<%=node.GetAttribute("name")%>' /></td>


		<td align=center style="color:#090;font-weight:600;"><%=getType(node.GetAttribute("type"))%></td>
		<td align=center><%=node.GetAttribute("name")%></td>
		<td >
		<%=node.GetAttribute("intro")%>
		<br><font color="gray">更新时间间隔为：<font color="red"><%=node.GetAttribute("update")%></font>&nbsp;秒。</font>
		</td>
		<td style="color:#666;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:10px;padding:0 0 0 8px!important;"><%=node.GetAttribute("addTime")%><br><font color="red"><%=node.GetAttribute("updateTime")%></font></td>
		<td align=center style="color:#666;font-family:tahoma,Arial,Helvetica,sans-serif;font-size:10px;padding:0 0 0 8px!important;"><span style="font-size:12px;font-weight: 600;">admin</span><br><font color="gray"><%=node.GetAttribute("IP")%></font></td>
		<td align=center>
		<a href="#" onclick="document.myform.action.value='modify';document.myform.eName.value='<%=node.GetAttribute("name")%>';document.myform.submit();">编辑</a> &nbsp;<a href="../jsView.htm?action=<%=node.GetAttribute("name")%>" target="blank">预览</a>
		</td>
	</tr>
<%Next%>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30">
		<label for="chkAll"><input type="checkbox" id="chkAll" name="chkAll" onclick="CheckAll(this.form);"> 选中本页所有</label>
		<input type="hidden" name="action" value="del"><input type="hidden" name="eName" value=""><input type="submit" name="Submit" value="删除" onclick="return confirm('确定要删除选中的记录吗？');" >
	</td>
  </tr>
</table>
</form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
	<%
	End Sub
Sub add()
	If  action="modify" Then

		Set xmlDoc = New Cls_XmlDoc
		xmlDoc.Unicode = False

		If Not xmlDoc.LoadXml("../xmlData/jsTemplate.config") Then
			oblog.ShowMsg "模板文件不存在，无法完成操作",""
		End If
		Dim node
		Set node = XmlDoc.NodeObj("template[@name='"&eName&"']")

		Intro = XmlDoc.AtrributeValue("template[@name='"&eName&"']","intro")
		eType = XmlDoc.AtrributeValue("template[@name='"&eName&"']","type")
		Update = XmlDoc.AtrributeValue("template[@name='"&eName&"']","update")
		order = XmlDoc.AtrributeValue("template[@name='"&eName&"']","order")

		head =  node.selectSingleNode("head").text
		skinmain =  node.selectSingleNode("main").text
		foot =  node.selectSingleNode("foot").text
	End if
	%>
<style type="text/css">
.main_content_leftbg div ul label { width: 200px; text-align: right; }
fieldset legend { font-weight: 600; }
#skin_info ol { padding: 0 0 0 12px; margin: 0 0 0 20px; }
#skin_info ol li { list-style: disc inline none; }

</style>

<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">添加JS调用</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="post" action="?action=saveadd" name="TheForm">
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td width="30%" class="td2" align="right">名称：</td>
		<td width="70%" class="td1"><input type="text" name="eName" size="20" Maxlength="10" onkeyup="OutputNewsCode(this.value);" value="<%=eName%>" <%If  action="modify" Then Response.Write "readonly" %> >(英文或者数字)</td>
	</tr>
	<tr>
		<td class="td2" align="right">调用代码：</td>
		<td class="td1"><input type="text" name="code" id="code" style="width: 100%;" size="60" disabled value="<script src=&quot;<%=Trim(oblog.CacheConfig(3))%>jsNew.asp?action=<%=eName%>&quot;></script>"></td>
	</tr>
	<tr>
		<td class="td2" align="right">调用说明：</td>
		<td class="td1"><input type="text" name="Intro" size="30" Maxlength="30" value="<%=Intro%>"></td>
	</tr>
	<tr>
		<td class="td2" align="right">调用类型：</td>
		<td class="td1">
			<select NAME="eType" ID="eType" onchange="NewsTypeSel(this.selectedIndex)">
				<option value="0">请选择</option>
<%
Dim i
For i = 1 To UBound(getType)
%>
				<option value="<%=i%>" <%If Int(OB_IIF(eType,0)) = i Then Response.Write "selected"%>><%=getType(i)%></option>
<%Next%>
				</select>
		</td>
	</tr>
	<tr>
		<td class="td2" align="right">数据更新间隔：</td>
		<td class="td1"><input type="text" name="Update" value="<%=Update%>">(单位：秒)</td>
	</tr>
</table>
<div id="News"></div>

<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
<!-- 调用模板设置 -->
	<tr>
      <td height="25" colspan="2" class="topbg"><b>调用模板设置&nbsp;（请用HTML代码填写）</b></td>
    </tr>
	<tr>
		<td width="30%" class="td2" align="right" valign="top">模板开始标记部分</td>
		<td width="70%" class="td2"><textarea name="head" ID="head" style="width:100%;" rows="3"><%=head%></textarea></td>
	</tr>
	<tr>
		<td class="td2" align="right" valign="top">模板主体循环标记部分
			<fieldset title="模板变量">
				<legend>&nbsp;模板变量说明&nbsp;</legend>
				<div id="skin_info" align="left">请选择调用类型。</div>
			</fieldset>
		</td>
		<td class="td2" valign="top">
			<div id="DisInput"></div>
			<textarea name="main" ID="main" style="width:100%;" rows="10"><%=skinmain%></textarea>
		</td>
	</tr>
	<tr>
		<td class="td2" align="right" valign="top">模板结束标记部分</td>
		<td class="td2"><textarea name="foot" ID="foot" style="width:100%;" rows="3"><%=foot%></textarea></td>
	</tr>
<!-- 调用模板设置 -->
	<tr>
      <td height="40" colspan="2" align="center" class="tdbg" ><%If action="modify" Then %><input type="hidden" name="modify" value="1" /><%End if%><input type="submit" class="button" value=" 提交设置 "></td>
	</tr>
</table>

</form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>

<!-- 设置信息部分 -->
<div id="News_1" style="display:none">
<!-- 统计信息 -->
无设置信息
</div>
<div id="News_2" style="display:none">
<!-- 用户信息 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>用户信息设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn" value="<%
		If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","topN"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">博客名长度限制：</td>
		<td class="td1"><input type="text" id="length" name="length" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","length"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">排序方式：</td>
		<td class="td1">

			<select id="order" name="order">
				<option value="0" <%If order = "0" Then Response.Write "selected" %>>用户日志总数</option>
				<option value="1" <%If order = "1" Then Response.Write "selected" %>>用户访问量</option>
				<option value="2" <%If order = "2" Then Response.Write "selected" %>>用户积分</option>
				<option value="3" <%If order = "3" Then Response.Write "selected" %>>注册日期倒序</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="td2" align="right">是否推荐用户：</td>
		<td class="td1"><input type="checkbox" id="isbest" name="isbest" value="1" <%
		If  action="modify" Then
			if XmlDoc.AtrributeValue("template[@name='"&eName&"']","isbest") = "1" Then
				Response.Write "checked"
			End If
		End if

			%> /></td>
	</tr>
</table>
</div>
<div id="News_3" style="display:none">
<!-- 站点公告 -->
无设置信息
</div>
<div id="News_4" style="display:none">
<!-- 分类 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>系统分类设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">分类类型：</td>
		<td class="td1" width="70%">
			<select id="classType" name="classType">
				<option value="-1" <%If order = "-1" Then Response.Write "selected" %>>博客分类</option>
				<option value="0" <%If order = "0" Then Response.Write "selected" %>>日志分类</option>
				<option value="1" <%If order = "1" Then Response.Write "selected" %>>相册分类</option>
				<option value="2" <%If order = "2" Then Response.Write "selected" %>>圈子分类</option>
			</select>
		</td>
	</tr>
</table>
</div>
<div id="News_5" style="display:none">
<!-- 日志调用 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>日志调用设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","topN"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">标题长度限制：</td>
		<td class="td1"><input type="text" id="length" name="length" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","length"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">日志分类：</td>
		<td class="td1"><select id="classid" name="classid">
		<% Dim classid
		If  action="modify" Then
			classid =  XmlDoc.AtrributeValue("template[@name='"&eName&"']","classid")
		Else
			classid = 0
		End if
		%>
		<%=oblog.show_class(2,classid,0)%></select></td>
	</tr>
	<tr>
		<td class="td2" align="right">用户ID：</td>
		<td class="td1"><input type="text" id="userid" name="userid" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","userid"))%>" size="10" />（如果要调用某一个人的日志，请指定已经存在的某个用户的ID）</td>
	</tr>
	<tr>
		<td class="td2" align="right">日志时间范围：</td>
		<td class="td1"><input type="text" id="sdate" name="sdate" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","sdate"))%>" size="10" />天内</td>
	</tr>
	<tr>
		<td class="td2" align="right">排序方式：</td>
		<td class="td1">
			<select id="order" name="order">
				<option value="0" <%If order = "0" Then Response.Write "selected" %>>日志访问数</option>
				<option value="1" <%If order = "1" Then Response.Write "selected" %>>日志回复数</option>
				<option value="2" <%If order = "2" Then Response.Write "selected" %>>发布日期倒序</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="td2" align="right">是否显示系统分类：</td>
		<td class="td1"><input type="checkbox" id="iscid" name="iscid" value="1" <%
		If  action="modify" Then
			if XmlDoc.AtrributeValue("template[@name='"&eName&"']","isClass") = "1" Then
				Response.Write "checked"
			End If
		End if%> /></td>
	</tr>
	<tr>
		<td class="td2" align="right">是否显示用户专题：</td>
		<td class="td1"><input type="checkbox" id="issid" name="issid" value="1" <%
		If  action="modify" Then
			if XmlDoc.AtrributeValue("template[@name='"&eName&"']","isSubject") = "1" Then
				Response.Write "checked"
			End If
		End if%> /></td>
	</tr>
	<tr>
		<td class="td2" align="right">是否精华日志：</td>
		<td class="td1"><input type="checkbox" id="isbest" name="isbest" value="1"<%
		If  action="modify" Then
			if XmlDoc.AtrributeValue("template[@name='"&eName&"']","isbest") = "1" Then
				Response.Write "checked"
			End If
		End if%> /></td>
	</tr>
	<tr>
		<td class="td2" align="right">时间格式：</td>
		<td class="td1">
		<%
		Dim formatTime
		If  action="modify" Then
			formatTime = XmlDoc.AtrributeValue("template[@name='"&eName&"']","formatTime")
		End If
		%>
			<select name="FormatTime" id="FormatTime">
				<option value="0" <%If formatTime = "0" Then Response.Write "selected" %>>YYYY-M-D H:M:S(长格式)</option>
				<option value="1" <%If formatTime = "1" Then Response.Write "selected" %>>YYYY年M月D</option>
				<option value="2" <%If formatTime = "2" Then Response.Write "selected" %>>YYYY-M-D</option>
				<option value="3" <%If formatTime = "3" Then Response.Write "selected" %>>H:M:S</option>
				<option value="4" <%If formatTime = "4" Then Response.Write "selected" %>>hh:mm</option>
			</select>（按服务器时间区域格式显示。）
		</td>
	</tr>
</table>
</div>
<div id="News_6" style="display:none">
<!-- 相片调用 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>相片调用设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","topN"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">每行显示记录数：</td>
		<td class="td1"><input type="text" id="br" name="br" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","br"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">排序方式：</td>
		<td class="td1">
			<select id="order" name="order">
				<option value="0" <%If order = "0" Then Response.Write "selected" %>>发布时间</option>
				<option value="1" <%If order = "1" Then Response.Write "selected" %>>访问数</option>
				<option value="2" <%If order = "2" Then Response.Write "selected" %>>评论数</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="td2" align="right">是否为调用相片（不选为相册）：</td>
		<td class="td1"><input type="checkbox" id="isalbum" name="isalbum" value="1" <%
		If  action="modify" Then
			if XmlDoc.AtrributeValue("template[@name='"&eName&"']","isalbum") = "1" Then
				Response.Write "checked"
			End If
		End if%> />是</td>
	</tr>
</table>
</div>
<div id="News_7" style="display:none">
<!-- 博客之星 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>博客之星调用设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn"   value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","topN"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">每行显示记录数：</td>
		<td class="td1"><input type="text" id="br" name="br" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","br"))%>" size="10" /></td>
	</tr>
</table>
</div>
<div id="News_8" style="display:none">
<!-- 圈子列表调用 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>圈子列表调用设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","topN"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">标题长度限制：</td>
		<td class="td1"><input type="text" id="len" name="len" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","length"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">排序方式：</td>
		<td class="td1">

			<select id="order" name="order">
				<option value="0" <%If order = "0" Then Response.Write "selected" %>>圈子成员人数倒序</option>
				<option value="1" <%If order = "1" Then Response.Write "selected" %>>圈子帖子数目倒序</option>
				<option value="2" <%If order = "2" Then Response.Write "selected" %>>创建日期倒序</option>
			</select>
		</td>
	</tr>
</table>
</div>
<div id="News_9" style="display:none">
<!-- 圈子日志调用 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>圈子日志调用设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","topN"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">圈子ID（多个请用|分开，比如1|2|3 ，可留空）：</td>
		<td class="td1"><input type="text" id="teamid" name="teamid" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","teamid"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">标题长度限制：</td>
		<td class="td1"><input type="text" id="length" name="length" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","length"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">时间格式：</td>
		<td class="td1">
			<select name="FormatTime" ID="FormatTime">
				<option value="0" <%If formatTime = "0" Then Response.Write "selected" %>>YYYY-M-D H:M:S(长格式)</option>
				<option value="1" <%If formatTime = "1" Then Response.Write "selected" %>>YYYY年M月D</option>
				<option value="2" <%If formatTime = "2" Then Response.Write "selected" %>>YYYY-M-D</option>
				<option value="3" <%If formatTime = "3" Then Response.Write "selected" %>>H:M:S</option>
				<option value="4" <%If formatTime = "4" Then Response.Write "selected" %>>hh:mm</option>
			</select>（按服务器时间区域格式显示）
		</td>
	</tr>
</table>
</div>
<div id="News_10" style="display:none">
<!-- 标签（TAG）调用 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>标签（TAG）调用设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","topN"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">每行显示记录数：</td>
		<td class="td1"><input type="text" id="br" name="br" value="5" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">排序方式：</td>
		<td class="td1">
			<select id="order" name="order">
				<option value="0" <%If order = "0" Then Response.Write "selected" %>>使用频度倒序</option>
				<option value="1" <%If order = "1" Then Response.Write "selected" %>>创建日期倒序</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="td2" align="right">表现形式：</td>
		<td class="td1">
		<%
		Dim iscloud
		If  action="modify" Then
			iscloud = XmlDoc.AtrributeValue("template[@name='"&eName&"']","iscloud")
		End If
		%>
			<select id="iscloud" name="iscloud">
				<option value="0" <%If iscloud = "0" Then Response.Write "selected" %>>列表</option>
				<option value="1" <%If iscloud = "1" Then Response.Write "selected" %>>云图</option>
			</select>
		</td>
	</tr>
</table>
</div>
<div id="News_11" style="display:none">
<!-- 用户推荐（DIGG）日志调用 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>用户推荐（DIGG）日志调用设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn" value="<%If  action="modify" Then  Response.Write (XmlDoc.AtrributeValue("template[@name='"&eName&"']","topN"))%>" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">排序方式：</td>
		<td class="td1">
			<select id="order" name="order">
				<option value="0" <%If order = "0" Then Response.Write "selected" %>>推荐次数倒序</option>
				<option value="1" <%If order = "1" Then Response.Write "selected" %>>创建日期倒序</option>
				<option value="2" <%If order = "2" Then Response.Write "selected" %>>最后推荐日期倒序</option>
			</select>
		</td>
	</tr>
</table>
</div>
<div id="News_12" style="display:none">
<!-- 被推荐（DIGG）用户信息调用 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>被推荐（DIGG）用户信息调用设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">调用条数：</td>
		<td class="td1" width="70%"><input type="text" id="topn" name="topn" value="10" size="10" /></td>
	</tr>
	<tr>
		<td class="td2" align="right">排序方式：</td>
		<td class="td1">
			<select id="order" name="order">
				<option value="0" <%If order = "0" Then Response.Write "selected" %>>推荐次数倒序</option>
				<option value="1" <%If order = "1" Then Response.Write "selected" %>>用户ID倒序</option>
			</select>
		</td>
	</tr>
</table>
</div>

<div id="News_13" style="display:none">
<!-- 登录窗口 -->
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<td height="25" colspan="2" class="topbg"><b>登录窗口设置</b></td>
    </tr>
	<tr>
		<td class="td2" width="30%" align="right">布局类型：</td>
		<td class="td1" width="70%">
			<input type="checkbox"  name = "login" id="login1" value="1" <%If order = "1" Then Response.Write "checked" %> /> 横向（默认竖向）
		</td>
	</tr>
</table>
</div>
<!-- 变量说明 -->
<div id="skininfo_0" style="display:none"></div>

<div id="skininfo_1" style="display:none">
	<ol>
		<li>博客总数：$usercount$</li>
		<li>日志总数：$logcount$</li>
		<li>评论总数：$commentcount$</li>
		<li>留言总数：$messagecount$</li>
		<li>今天日志：$logtoday$</li>
		<li>昨天日志：$logyestoday$</li>
	</ol>
</div>

<div id="skininfo_2" style="display:none">
	<ol>
		<li>用户博客地址：$userurl$</li>
		<li>用户名：$username$</li>
		<li>用户博客名：$blogname$</li>
		<li>用户日志总数：$logcount$</li>
	</ol>
</div>

<div id="skininfo_3" style="display:none">
	<ol>
		<li>公告内容：$placard$</li>
	</ol>
</div>

<div id="skininfo_4" style="display:none">
	<ol>
		<li>分类地址 ：$classurl$</li>
		<li>分类名 ：$classname$</li>
	</ol>
</div>

<div id="skininfo_5" style="display:none">
	<ol>
		<li>日志地址：$logurl$</li>
		<li>日志标题：$topic$</li>
		<li>用户博客地址：$userurl$</li>
		<li>日志作者：$postname$</li>
		<li>日志发布时间：$posttime$</li>
		<li>日志访问量：$iis$</li>
		<li>日志评论数：$commentnum$</li>
		<li>日志分类名：$classname$</li>
		<li>日志分类URL：$classurl$</li>
		<li>日志专题名：$subjectname$</li>
		<li>日志专题URL：$subjecturl$</li>
	</ol>
</div>

<div id="skininfo_6" style="display:none">
	<ol>
		<li>用户相册地址：$albumurl$</li>
		<li>相片地址：$imgsrc$</li>
		<li>相片介绍：$readme$</li>
		<li>换行标志：$br$（如果设定了每行显示记录，请在模板的主体循环标记部分最后加上$br$）</li>
	</ol>
</div>

<div id="skininfo_7" style="display:none">
	<ol>
		<li>博客之星地址：$userurl$</li>
		<li>用户博客地址：$blogurl$</li>
		<li>图片地址：$picurl$</li>
		<li>博客之星介绍：$info$</li>
		<li>博客博客名：$blogname$</li>
		<li>表格关闭标志：$tr$（如果设定了每行显示记录，请在模板的主体循环标记部分最后加上$tr$）</li>
	</ol>
</div>

<div id="skininfo_8" style="display:none">
	<ol>
		<li>圈子LOGO：$ico$</li>
		<li>圈子地址：$gurl$</li>
		<li>圈子名：$tname$</li>
		<li>圈子成员：$count0$</li>
		<li>圈子帖子总数：$count1$</li>
	</ol>
</div>

<div id="skininfo_9" style="display:none">
	<ol>
		<li>圈子日志地址：$posturl$</li>
		<li>帖子标题：$topic$</li>
		<li>帖子作者：$author$</li>
		<li>帖子发布时间：$addtime$</li>
	</ol>
</div>

<div id="skininfo_10" style="display:none">
	<ol>
		<li>TagUrl：$tagurl$</li>
		<li>Tag名：$tagname$</li>
		<li>Tag使用次数：$num$</li>
	</ol>
</div>

<div id="skininfo_11" style="display:none">
	<ol>
		<li>用户地址：$userurl$</li>
		<li>推荐次数：$num$</li>
		<li>日志地址：$url$</li>
		<li>日志标题：$title$</li>
		<li>添加时间：$addtime$</li>
	</ol>
</div>
<div id="skininfo_12" style="display:none">
	<ol>
		<li>用户地址：$userurl$</li>
		<li>推荐次数：$num$</li>
		<li>用户头像：$imgsrc$</li>
		<li>用户名：$username$</li>
	</ol>
</div>
<div id="skininfo_13" style="display:none">
	<ol>
		<li><strong><font color="red">请勿更改模板部分的内容</font></strong></li>
	</ol>
</div>
<%End Sub

Sub delNode()
	Dim eName,arrTemp,nodeTemp,xmlTemp
	Dim i
	Set xmlTemp = new Cls_xmlDoc
	xmlTemp.Unicode = False
	If Not xmlTemp.LoadXml("../xmlData/jsTemplate.config") Then
		Response.Write  "模板文件不存在，无法完成操作"
		Response.End
	End If
	eName = Trim(request("ename"))
	If eName ="" Then
		oblog.ShowMsg "请选择一个项目",""
	Else
		arrTemp = Split (eName,",")
		For Each eName In  Request("ename")
			Set nodeTemp = xmlTemp.NodeObj("template[@name='"&eName&"']")
			If Not (nodeTemp Is Nothing ) Then
				xmlTemp.removeChild("template[@name='"&eName&"']")
			End if
		Next
		xmlTemp.Save()
		oblog.ShowMsg "删除成功",""
	End if
End Sub

Sub saveAdd()

	If eName = "" Then
		oblog.ShowMsg "名称不能为空","back"
	End If
	If Intro = "" Then
		oblog.ShowMsg "调用说明不能为空","back"
	End If
	If eType = 0 Then
		oblog.ShowMsg "请选择一个调用类别","back"
	End If
	If Update ="" Or Not IsNumeric(Update) Then Update = 600
	Set xmlDoc = new Cls_xmlDoc
	xmlDoc.Unicode = False

	If Not xmlDoc.LoadXml("../xmlData/jsTemplate.config") Then
		oblog.ShowMsg "模板文件不存在，无法完成操作","back"
	End If

	Set node = XmlDoc.NodeObj("template[@name='"&eName&"']")

	If isModify = "" Then
	'非编辑模式不允许重名
		If Not (node Is Nothing ) Then
			oblog.ShowMsg "此名称已经存在，请换名重试","back"
		End If
		XmlDoc.InsertElement2 XmlDoc.NodeObj("root"),"template","",False,"name",eName
	End If

	Set node = XmlDoc.NodeObj("template[@name='"&eName&"']")

	xmlDoc.setAttributeNode "template[@name='"&eName&"']","type",etype
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","intro",intro
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","update",update
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","IP",oblog.UserIp
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","admin",session("adminname")

	If isModify = "" Then
		xmlDoc.setAttributeNode "template[@name='"&eName&"']","addTime",Now()
		XmlDoc.InsertElement node,"head",head,False,True
		XmlDoc.InsertElement node,"main",skinmain,False,True
		XmlDoc.InsertElement node,"foot",foot,False,True
	Else
		XmlDoc.UpdateNodeText2 node.selectSingleNode("head"),head,True
		XmlDoc.UpdateNodeText2 node.selectSingleNode("main"),skinmain,True
		XmlDoc.UpdateNodeText2 node.selectSingleNode("foot"),foot,True
	End If

	xmlDoc.setAttributeNode "template[@name='"&eName&"']","updateTime",Now()

	Select Case eType
		Case 1
		Case 2 : Call listUser()
		Case 3
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
	End Select


	If isModify = "" Then
		XmlDoc.InsertElement node,"sql",sql,True,True
	Else
		XmlDoc.UpdateNodeText2 node.selectSingleNode("sql"),sql,True
	End if
	XmlDoc.save()
	oblog.ShowMsg "操作成功",""
End Sub

Sub listUser()

	xmlDoc.setAttributeNode "template[@name='"&eName&"']","length",OB_IIF(length,20)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(order,0)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","isbest",OB_IIF(isbest,0)

	Select Case CLng(order)
		Case 0:order="log_count DESC,userid"
		Case 1:order="user_siterefu_num DESC,userid"
		Case 2:order="scores DESC,userid"
		Case 3:order="userid"
	End Select

	If isbest = "1" Then
		Sql = "SELECT TOP "&topN&" username,log_count,blogname,userid,user_domain,user_domainroot FROM [oblog_user] WHERE user_isbest=1 and (is_log_default_hidden=0 or is_log_default_hidden is null) ORDER BY log_count,userid DESC"
	Else
		Sql = "SELECT TOP "&topN&" username,log_count,blogname,userid,user_domain,user_domainroot FROM [oblog_user] where (is_log_default_hidden=0 or is_log_default_hidden is null) ORDER BY "&order&" DESC"
	End If

End Sub

Sub listClass()

	Dim classType
	classType = Trim(request("classType"))

	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(classType,0)

	If classType = "-1" Then
		Sql = "SELECT id,classname FROM [oblog_userclass] ORDER BY RootID,OrderID"
	Else
		Sql = "SELECT id,classname FROM  [oblog_logclass] WHERE idtype= "&CLng(classType)&" ORDER BY RootID,OrderID"
	End if
End Sub

Sub showLog()

	Dim isClass,isSubject
	Dim classid
	If Trim(request("iscid")) = "1" Then
		isClass = 1
	Else
		isClass = 0
	End If
	If Trim(request("issid")) = "1" Then
		isSubject = 1
	Else
		isSubject = 0
	End if
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","length",OB_IIF(length,20)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","formatTime",formatTime
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","isClass",isClass
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","isSubject",isSubject
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","classid",OB_IIF(Trim(request("classid")),0)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","userid",OB_IIF(Trim(request("userid")),0)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","sdate",OB_IIF(Trim(request("sdate")),7)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(order,0)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","isbest",isbest

	Dim usersql,isbestsql,sdatesql,classsql
	Dim sDate
	sDate = Request("sdate")
	if Trim(request("userid"))>"0" then
		usersql=" AND a.userid="&CLng(request("userid"))
	Else
		usersql=""
	End If
	If Not IsNumeric(sDate) then
		oblog.showMsg ("错误的时间参数"),"back"
	end If
	If isbest = "1" Then
		isbestsql=" AND isbest=1"
	Else
		isbestsql = ""
	End If
	If Is_Sqldata = 0 Then
		sdatesql = sdatesql&" DATEDIFF("&G_Sql_d&",a.truetime,"&G_Sql_Now&")<"&Int(sdate)&" "
	Else
		sdate = DateAdd("d",-1*Abs(sdate),Now())
		sdate = GetDateCode(sdate,0)
		sdatesql = sdatesql&" truetime>'"&sdate&"'"
	End If
	classid = Trim(request("classid"))
	If classid  = "0" Then
		classsql = ""
	Else
		Dim rs,ustr
		set rs=oblog.execute("SELECT id FROM oblog_logclass WHERE parentpath LIKE '"&CLng(classid)&",%' OR parentpath LIKE '%,"&CLng(classid)&"' OR parentpath LIKE '%,"&CLng(classid)&",%'")
		While Not rs.EOF
			ustr=ustr&","&rs(0)
			rs.MoveNext
		Wend
		ustr=classid&ustr
		classsql=" AND classid IN ("&ustr&")"
	End If

	Select Case CLng(order)
		Case 0:order="iis DESC,logid"
		Case 1:order="commentnum DESC,logid"
		Case 2:order="logid"
	End Select

	Sql = "SELECT TOP "&topN&" author,topic,logid,classid,subjectid,truetime,iis,commentnum,a.userid,user_domain,user_domainroot FROM oblog_log a INNER JOIN oblog_user b ON B.userid=A.userid WHERE "&sdatesql&usersql&isbestsql&" AND passcheck=1 AND a.isdel=0 AND isdraft=0 AND  (IsSpecial = 0 OR IsSpecial IS NULL) and (b.is_log_default_hidden=0 or b.is_log_default_hidden is null) "&classsql&" ORDER BY "&order&" DESC"
End Sub

Sub showPhoto()

	xmlDoc.setAttributeNode "template[@name='"&eName&"']","br",OB_IIF(Trim(request("br")),1)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(order,0)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","isalbum",OB_IIF(Trim(request("isalbum")),0)

	If Trim(request("isalbum")) = "1" Then
		Select Case CLng(order)
			Case 0:order="photoID "
			Case 1:order="views DESC,photoID"
			Case 2:order="commentnum DESC,photoID"
		End Select
	Else
		Select Case CLng(order)
			Case 0:order="subjectid "
			Case 1:order="views DESC,subjectid"
		End Select
	End If
	If Trim(request("isalbum")) = "1" Then
		Sql = "SELECT TOP "&topN&"  photo_path,photo_readme,userid FROM oblog_album  where (ishide = 0 OR ishide IS NULL) ORDER BY "&ORDER&" DESC"
	else
		Sql = "SELECT TOP "&topN&" photo_path,subjectname,userid,subjectid,subjectlognum FROM oblog_subject WHERE subjecttype = 1 AND (ishide = 0 OR ishide IS NULL) ORDER BY "&ORDER&" DESC "
	End if
End Sub

Sub showBlogStar()

	xmlDoc.setAttributeNode "template[@name='"&eName&"']","br",OB_IIF(Trim(request("br")),1)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)

	Sql = "SELECT TOP "&topN&" userurl , picurl ,info ,blogname,userid FROM oblog_blogstar WHERE ispass=1 ORDER BY ID DESC"
End Sub

Sub showTeam()
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","length",OB_IIF(length,20)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(order,0)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","islogo",Trim(request("islogo"))

	Select Case CLng(order)
		Case 0:order="icount0 DESC ,teamid "
		Case 1:order="(icount1+icount2) DESC ,teamid"
		Case 2:order="teamid"
	End Select
	Dim isbestsql
	If isbest = "1" Then
		isbestsql=" AND isbest=1"
	Else
		isbestsql = ""
	End If

	Sql = "SELECT TOP "&topN&" teamid,t_name,t_ico,icount0,(icount1+icount2) FROM oblog_team WHERE istate=3 AND isdel=0 "&isbestsql&" ORDER BY "&order&" DESC"
End Sub

Sub showTeamPost()

	xmlDoc.setAttributeNode "template[@name='"&eName&"']","length",OB_IIF(length,20)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","formatTime",formatTime
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","teamid",OB_IIF(Trim(request("teamid")),0)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","isuname",Trim(request("isuname"))
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","istime",Trim(request("istime"))

	Dim teamid,tsql
	teamid = Trim(request("teamid"))
	If teamid<>"" And teamid<>"0" Then
		teamid=Replace(teamid,"|",",")
		teamid  = FilterIDs(teamid)
		If teamid <> "" Then
			tsql =  " And teamid In (" & teamid & ") "
		Else
			tsql = ""
		End if
	End If
	Sql = "SELECT TOP "&topN&" teamid,postid,topic,addtime,author,userid FROM oblog_teampost WHERE 1=1 "&tsql&" AND idepth=0 AND isdel=0  ORDER BY postid DESC"
End Sub

Sub showTag()
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","br",OB_IIF(Trim(request("br")),5)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(order,0)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","iscloud",OB_IIF(Trim(request("iscloud")),0)

	Dim iscloud
	Dim ordersql
	iscloud = Trim (request("iscloud"))
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","iscloud",iscloud
	If iscloud=0 Then
		If order = "0" Then
			ordersql= " Order By iNum Desc,tagid DESC "
		Else
			ordersql= " Order By tagid DESC "
		End if
	Else
		If Is_Sqldata > 0 Then
			ordersql= " Order By Newid()"
		Else
			Randomize
			ordersql= " Order By Rnd(-(TagID+"&Rnd()&"))"
		End If
	End If
	'先取最近的N条，防止取出N年前的记录
	'然后对此N条记录进行排序等筛选操作
	Sql = "SELECT * FROM (SELECT TOP "&topN&" tagid,name,inum,iState FROM Oblog_Tags WHERE iNum>0 AND iState=1 ORDER BY tagid DESC) AS T  "&ordersql
End Sub

Sub showDigg()
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(order,0)

	Select Case CLng(order)
		Case 0:order="diggnum DESC,DiggID"
		Case 1:order="DiggID "
		Case 2:order="lastdiggtime "
	End Select
	Sql = "SELECT TOP "&topN&" diggnum,diggurl,diggtitle,addtime,author,authorid FROM oblog_userdigg WHERE istate = 1 ORDER BY "&order&" DESC"
End Sub

Sub showUserDigg()
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","topN",OB_IIF(topN,10)
	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(order,0)

	Select Case CLng(order)
		Case 0:order="diggs DESC,userid"
		Case 1:order="userid "
	End Select
	Sql = "SELECT TOP "&topN&" userid,User_Icon1,username,nickname,diggs FROM oblog_user WHERE lockuser=0 AND isdel=0 AND (is_log_default_hidden=0 or is_log_default_hidden is null) ORDER BY "&order&" DESC"
End Sub

Sub showLogin()

	xmlDoc.setAttributeNode "template[@name='"&eName&"']","order",OB_IIF(Trim(Request("login")),0)


End Sub
%>
	</html>
<script>
	function CheckAll(form)
	{
	  for (var i=0;i<form.elements.length;i++)
		{
		var e = form.elements[i];
		if (e.Name != "chkAll")
		   e.checked = form.chkAll.checked;
		}
	}
	function NewsTypeSel(index)
	{
		if (index > 0)
		{
		document.getElementById('skin_info').innerHTML = document.getElementById('skininfo_'+index).innerHTML;
		document.getElementById('News').innerHTML = document.getElementById('News_'+index).innerHTML;
		}

	}
	function OutputNewsCode(values)
	{
		document.getElementById('code').value='<scr'+'ipt src="<%=Trim(oblog.CacheConfig(3))%>jsNew.asp?action='+values+'"></scr'+'ipt>';
	}
</script>
<%
If action = "modify" Then
'修改模式下，控制隐藏层的显示
%>
<script>NewsTypeSel('<%=etype%>');</script>
<%End if%>