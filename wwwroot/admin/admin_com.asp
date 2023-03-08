<!--#include file="inc/inc_sys.asp"-->
<%
Const C_Items=23
Dim action
action=Trim(Request("action"))
Dim aObjects(12)
aObjects(0) = "Scripting.FileSystemObject"
aObjects(1) = "adodb.connection"
'-----------------------
aObjects(2) = "JMail.Message"				'JMail 4.3
aObjects(3) = "CDONTS.NewMail"				'CDONTS
aObjects(4) = "Persits.MailSender"			'ASPEMAIL
'-----------------------
aObjects(5) = "Adodb.Stream"				'Adodb.Stream
aObjects(6) = "Persits.Upload"				'Aspupload3.0
aObjects(7) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
aObjects(8) = "DvFile.Upload"				'DvFile-Up V1.0
aObjects(9) = "LyfUpload.UploadFile"
'-----------------------
aObjects(10) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
aObjects(11) = "Persits.Jpeg"				'AspJpeg
aObjects(12) = "sjCatSoft.Thumbnail"		'sjCatSoft.Thumbnail V2.6

if action="saveconfig" then
	call saveconfig()
else
	call showconfig()
end if


Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

sub showconfig()
dim rs,ac
set rs=oblog.execute("select * From oblog_config Where id=2")
ac=Split(rs("ob_value"),"$$")
set rs=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>站点配置</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">网 站 组 件 配 置</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="#SiteInfo">组件检测</a> | <a href="#SiteOption">名称定义</a>
      | <a href="#user">上传(图片)组件</a> | <a href="#show">邮件组件</a>
    </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">服 务 器 信 息</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">

<form method="POST" action="admin_com.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <td width="348" class="tdbg" height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td class="tdbg">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td width="348" class="tdbg" height=23>站点物理路径：<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td class="tdbg">数据库地址：</td>
  </tr>
  <tr>
    <td class="tdbg" height=23>FSO文本读写：
      <%
      If IsObjInstalled(aObjects(0))=false Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
    <td class="tdbg">数据库使用：
      <%If Not IsObjInstalled(aObjects(1)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
  </tr>
  <tr>
    <td class="tdbg" height=23>Jmail组件支持：
      <%If Not IsObjInstalled(aObjects(2)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
    <td class="tdbg">CDONTS组件支持：
      <%If Not IsObjInstalled(aObjects(3)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
  </tr>
    <tr>
      <td height="25" colspan="2" class="topbg"><a name="SiteOption"></a><b>基础组件设置</b></td>
    </tr>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >系统必须使用Scripting.FileSystemObject和Adodb.Stream组件</td>
      <td height="25" > 如果您的服务器上这两个组件的名称已经更改,请修改组件的名称</td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >Scripting.FileSystemObject：</td>
      <td height="25" ><%
      	If ac(1)="" Then
      		Call EchoInput("a1",40,100,"Scripting.FileSystemObject")
      	Else
      		Call EchoInput("a1",40,100,ac(1))
    	End If
      	%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >Adodb.Stream：</td>
      <td height="25" >
      	<%
      	If ac(1)="" Then
      		Call EchoInput("a2",40,100,"Adodb.Stream")
      	Else
      		Call EchoInput("a2",40,100,ac(2))
    	End If
      	%></td>
    </tr>
     <tr>
      <td height="25" colspan="2" class="topbg"><a name="show"></a><b>邮件选项</b></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25">说明</td>
      <td height="25">启用该参数后，将可以开启邮件有效性检测等环节。但是邮件发送对服务器资源消耗比较大，请慎重选择。
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >是否启用邮件注册验证机制</td>
      <td height="25" ><% Call EchoRadio("a3","","",ac(3))%></td>
    </tr>
<!--     <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >请选择您要使用的邮件选项</td>
      <td height="25" ><%=MakeSelect_Mail(ac(4))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" ><p>邮件地址</p></td>
      <td height="25" ><% Call EchoInput("a5",40,50,ac(5))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >SMTP主机</td>
      <td height="25" ><% Call EchoInput("a6",40,50,ac(6))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" ><p>用户名</p></td>
      <td height="25" ><% Call EchoInput("a7",40,50,ac(7))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >口令</td>
      <td height="25" ><% Call EchoInput("a8",40,50,ac(8))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >是否启用SMTP验证</td>
       <td height="25" ><% Call EchoRadio("a9","","",ac(9))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td height="25" >邮件发送队列数目(每次最大10个,每个目标地址算一个)</td>
      <td height="25" ><% Call EchoInput("a10",40,50,ac(10))%></td>
    </tr> -->
    <tr>
      <td height="25" colspan="2" class="topbg"><a name="upload" id="user"></a><strong>上传选项</strong></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" ><p>选取上传组件：<br>
          (可以到<a href="http://www.oblog.cn" target="_blank">http://www.oblog.cn</a>下载Aspupload3.0组件)
        </p>
        </td>
      <td height="25" ><%=MakeSelect_Upload(ac(11))%>
      	</td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >图片缩略图及水印设置开关：<br>
        (服务器需安装AspJpeg组件，可到<a href="http://www.oblog.cn" target="_blank">http://www.oblog.cn</a>下载)
      </td>
      <td height="25" ><%=MakeSelect_Photo(ac(12))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >上传图片添加水印文字信息（可为空或0）:</td>
      <td height="25" ><% Call EchoInput("a13",40,50,ac(13))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >上传添加水印字体大小:</td>
      <td height="25" ><% Call EchoInput("a14",40,50,ac(14))%><b>px</b> </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >上传添加水印字体颜色:</td>
      <td height="25" ><input type="text" name="a15" ID="d_a15" size=10 value="<%if ac(15)="" then Response.Write("#FFFFFF") else Response.Write(ac(15))%>">
        <img border=0 id="s_a15" src="images/rect.gif" style="cursor:pointer;background-Color:<%=ac(15)%>;" onclick="SelectColor('a15');" title="选取颜色!">
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >上传添加水印字体名称:</td>
      <td height="25" ><select name="a16" id="a16">
          <option value="宋体" <%If ac(16)="宋体" Then Response.Write "selected"%>>宋体</option>
          <option value="楷体_GB2312" <%If ac(16)="楷体_GB2312" Then Response.Write "selected"%>>楷体</option>
          <OPTION value="Andale Mono" <%If ac(16)="Andale Mono" Then Response.Write "selected"%>>Andale Mono</OPTION>
          <OPTION value="Arial" <%If ac(16)="Arial" Then Response.Write "selected"%>>Arial</OPTION>
          <OPTION value="Arial Black" <%If ac(16)="Arial Black" Then Response.Write "selected"%>>Arial Black</OPTION>
          <OPTION value="Century Gothic"<%If ac(16)="Century Gothic" Then Response.Write "selected"%>>Century Gothic</OPTION>
          <OPTION value="Comic Sans MS" <%If ac(16)="Comic Sans MS" Then Response.Write "selected"%>>Comic Sans MS</OPTION>
          <OPTION value="Courier New" <%If ac(16)="Courier New" Then Response.Write "selected"%>>Courier New</OPTION>
          <OPTION value="Georgia" <%If ac(16)="Georgia" Then Response.Write "selected"%>>Georgia</OPTION>
          <OPTION value="Impact" <%If ac(16)="Impact" Then Response.Write "selected"%>>Impact</OPTION>
          <OPTION value="Tahoma" <%If ac(16)="Tahoma" Then Response.Write "selected"%>>Tahoma</OPTION>
          <OPTION value="Times New Roman" <%If ac(16)="Times New Roman" Then Response.Write "selected"%>>Times New Roman</OPTION>
          <OPTION value="Stencil" <%If ac(16)="Stencil" Then Response.Write "selected"%>>Stencil</OPTION>
          <OPTION value="Verdana" <%If ac(16)="Verdana" Then Response.Write "selected"%>>Verdana</OPTION>
        </select></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >上传水印字体是否粗体:</td>
      <td height="25" ><% Call EchoRadio("a17","","",ac(17))%>
        </select></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >上传图片添加水印LOGO图片信息（可为空或0）:<br>
        填写LOGO的图片相对路径
        <br/>
        相对根目录的路径，如放置于根目录下，则直接写路径 test.jpg
         <br/>
        如放置于根目录下的images目录下，则为 images/test.jpg</td>
      <td height="25" ><% Call EchoInput("a18",40,50,ac(18))%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >上传图片添加水印透明度:</td>
      <td height="25"><% Call EchoInput("a19",40,50,ac(19))%>如60%请填写0.6 </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >水印图片去除底色:<br>
        保留为空则水印图片不去除底色。</td>
      <td height="25" ><input type="text" name="a20" ID="d_a20" size=10 value="<%if ac(20)="" then Response.Write("#FFFFFF") else Response.Write(ac(20))%>">
        <img border=0 id="s_a20" src="images/rect.gif" style="cursor:pointer;background-Color:<%=ac(20)%>;" onclick="SelectColor('a20');" title="选取颜色!">
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >水印文字或图片的长宽区域定义:<br>
        如水印图片的宽度和高度。</td>
      <td height="25" >高度：<% Call EchoInput("a21",5,5,ac(21))%>象素,
      	宽度：<% Call EchoInput("a22",5,5,ac(22))%>象素 </td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >上传图片添加水印LOGO位置坐标 :</td>
      <td height="25" ><select NAME="a23" id="upset_DrawXYType">
          <option value="0" <%if ac(23)=0 then Response.Write("selected")%>>左上</option>
          <option value="1" <%if ac(23)=1 then Response.Write("selected")%>>左下</option>
          <option value="2" <%if ac(23)=2 then Response.Write("selected")%>>居中</option>
          <option value="3" <%if ac(23)=3 then Response.Write("selected")%>>右上</option>
          <option value="4" <%if ac(23)=4 then Response.Write("selected")%>>右下</option>
        </select></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >&nbsp;</td>
      <td height="25" >&nbsp; </td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg" > <input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " > </td>
    </tr>
  </table>

</form>
<script language="javascript">

function SelectColor(what){
	var dEL = document.all("d_"+what);
	var sEL = document.all("s_"+what);
	var arr = showModalDialog("../images/selcolor.html", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
	if (arr) {
		dEL.value=arr;
		sEL.style.backgroundColor=arr;
	}
}
</script>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</body>
</html>
<%

end sub

'上传组件检测并生成选择表单
Function MakeSelect_Upload(sValue)
	Dim sRet,s1,s2,s3,s4
	select Case sValue
		Case "0"
			s2=" selected"
		Case "1"
			s3=" selected"
		Case "2"
			s4=" selected"
		Case Else
			s1=" selected"
	End select
	sRet="<select name=a11>" & vbcrlf
	sRet= sRet & "<option value=""999""" & s1 & ">关闭</option>" & vbcrlf
	If IsObjInstalled(aObjects(5)) Then sRet= sRet & "<option value=""0""" & s2 & ">无组件上传</option>" & vbcrlf
	If IsObjInstalled(aObjects(6)) Then sRet= sRet & "<option value=""1""" & s3 & ">Aspupload3.0组件 </option>" & vbcrlf
	If IsObjInstalled(aObjects(7)) Then sRet= sRet & "<option value=""2""" & s4 & ">SA-FileUp 4.0组件</option>" & vbcrlf
	sRet= sRet & "</select>"
	MakeSelect_Upload=sRet
	sRet=""
End Function

'自动检测缩略图组件并生成表单
Function MakeSelect_Photo(sValue)
	Dim sRet,sRet1,bRet,s1,s2,s3
	select Case sValue
		Case "0"
			s1=" selected"
		Case "1"
			s2=" selected"
		Case "2"
			s3=" selected"
		Case Else
			s1=" selected"
	End select
	If IsObjInstalled("Persits.Jpeg") Then
		bRet=true
		sRet1= "AspJpeg组件<font color=red><b>√</b>服务器支持!</font>"
	Else
		bRet=false
		sRet1= "AspJpeg组件<b>×</b>服务器不支持!"
	End If
	sRet= "<select name=a12>" & vbcrlf
	sRet= sRet & "<option value=""0""" & s1 & ">关闭缩略图及水印效果</option>" & vbcrlf
	If bRet Then
		sRet= sRet & "<option value=""1""" & s2 & ">开启缩略图及水印文字效果(推荐)</option>" & vbcrlf
		sRet= sRet & "<option value=""2""" & s3 & ">开启缩略图及水印图片效果</option>" & vbcrlf
	End If
	sRet= sRet & "</select>&nbsp;&nbsp;(" & sRet1 & ")"
	MakeSelect_Photo=sRet
	sRet=""
	sRet1=""
End Function

'自动检测邮件组件并生成选择表单
Function MakeSelect_Mail(sValue)
	Dim sRet,s1,s2,s3,s4
	select Case sValue
		Case "0"
			s2=" selected"
		Case "1"
			s3=" selected"
		Case "2"
			s4=" selected"
		Case Else
			s1=" selected"
	End select
	sRet="<select name=a4>" & vbcrlf
	sRet= sRet & "<option value=""999""" & s1 & ">关闭或不支持</option>" & vbcrlf
	If IsObjInstalled(aObjects(2)) Then sRet= sRet & "<option value=""0""" & s2 & ">JMail组件</option>" & vbcrlf
	If IsObjInstalled(aObjects(3)) Then sRet= sRet & "<option value=""1""" & s3 & ">CDONT(2000/2003自带)</option>" & vbcrlf
	If IsObjInstalled(aObjects(4)) Then sRet= sRet & "<option value=""2""" & s4 & ">AspMail组件</option>" & vbcrlf
	sRet= sRet & "</select>"
	MakeSelect_Mail=sRet
	sRet=""
End Function

Sub Saveconfig()
	If Request.QueryString <>"" Then Exit Sub
	Dim rs, i,sOpt
  	sOpt=""
  	For i=1 To C_Items
  		sOpt=sOpt & "$$" & Replace(Trim(Request.Form("a"&i)),"$","")
  	Next
  	sOpt=Now&sOpt
    If Not IsObject(conn) Then link_database
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open  "select * From oblog_config Where Id=2",conn,1,3
    If rs.Eof Then rs.AddNew
    rs("ob_value")=sOpt
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ReloadCache
	EventLog "进行修改服务器组件配置的操作!",""
    Set oblog=Nothing
    Response.Redirect "admin_com.asp"
End Sub

%>