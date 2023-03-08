<!--#include file="inc/inc_sys.asp"-->
<%
Dim  action,atype,sGuide,rs
action=Trim(Request("action"))
atype=lcase(trim(request("type")))
If atype="" Or IsNull(atype) Then atype="new" 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>防垃圾自定义验证问题模块配置</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">防垃圾自定义验证问题模块配置</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="?action=all">所有验证问题</a> | <a href="?action=new">新添加一个验证问题</a>  |  <a href="?action=p">批量生成验证问题</a>

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
Select Case LCase(action) 
	Case "modify","new","edit"
		Call Modify()
	Case "modifysave"
		Call SaveModify()
	Case "del"
		Call delone()
	Case "p"
		Call Batchcode()
	Case Else
		Call list()
End Select
%>
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

Function  Modify()
	Dim wen,da,vid
	If atype="new" And request("id")<>"" And IsNumeric(request("id")) Then 
		Set rs=oblog.execute("select * from oblog_Verifiydata where id="&int(trim(request("id"))))
		vid=rs("id")
		wen=rs("ask")
		da=rs("answer")
		Set rs=Nothing 
	End If 
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">新添加一个验证问题</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">

<form method="POST" action="admin_ask.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
 
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >验证问题提示：<br/>(例:请填入一个10以内的单数.(一位数字))</td>
      <td height="25" ><%Call EchoInput("wen",40,40,wen)%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="25" >验证问题答案：<br/>(例:  1|3|5|7|9  )</td>
      <td height="25" ><%Call EchoInput("da",40,20,da)%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''">
      <td height="30" colspan="2">(问题答案多个的话用 | 隔开,答案最好不要超过5个.每个答案不能超过5个汉字或10个数字或英文.可以数字与汉字混合.数字答案请注意半角全角最好都添上.您也可以直接给出答案提示.)</td>

    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg" > <input name="Action" type="hidden" id="Action" value="modifysave">
	  <input name="id" type="hidden" id="id" value="<%=vid%>">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " > </td>
    </tr>
  </table>

</form>
<%
End Function 
Sub SaveModify()
	If Request.QueryString <>"" Then Exit Sub
	Dim id
	id=trim(request("id"))
	If Not (id<>"" or IsNumeric(id)) Then id=-1
	id=Int (id)
    If Not IsObject(conn) Then link_database
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open  "select * From oblog_Verifiydata Where id="&id,conn,1,3
    If rs.Eof Then rs.AddNew
   	rs("ask")=Trim(request("wen"))
	rs("answer")=Trim(request("da"))
    rs.Update
    rs.Close
    Set rs = Nothing
    
	EventLog "添加修改了一条自定义验证问题.问题("&Trim(request("wen"))&"),答案("&Trim(request("da"))&") ",""
    Set oblog=Nothing
    Response.Redirect "admin_ask.asp"
End Sub
Sub delone()
If request("id")<>"" And IsNumeric(Trim(request("id"))) Then 
oblog.execute("delete from oblog_Verifiydata where id="&int(trim(request("id"))))
EventLog "删除了一条自定义验证问题.id为"&int(trim(request("id"))),""
oblog.ShowMsg "操作成功!", ""

Else
oblog.ShowMsg "没有要操作的id.", ""
End If 
End Sub 
Sub Batchcode()
oblog.ShowMsg "此功能稍后升级", ""
'批量操作随机生成验证问题类型为:
'两位数以内整数加减运算验证问题
'一位数乘法运算 
'两位以内字母排序
'by 蓝色 2007年6月29日

End Sub 

Sub list()
	sGuide="所有自定义验证问题列表 "
	if Request("page")<>"" then
    G_P_This=cint(Request("page"))
else
	G_P_This=1
end If
	G_P_FileName ="admin_ask.asp"
	set rs=Server.CreateObject ("Adodb.recordset")
	rs.open "select * from oblog_Verifiydata order by id desc",conn,1,1
	If rs.eof Or rs.bof Then
		Response.write "暂无数据!"
	Else
    G_P_AllRecords = rs.recordcount
        sGuide = sGuide & "(<font color=red>" & G_P_AllRecords & "</font>)"
        If G_P_This < 1 Then
            G_P_This = 1
        End If
        If (G_P_This - 1) * G_P_PerMax > G_P_AllRecords Then
            If (G_P_AllRecords Mod G_P_PerMax) = 0 Then
                G_P_This = G_P_AllRecords \ G_P_PerMax
            Else
                G_P_This = G_P_AllRecords \ G_P_PerMax + 1
            End If

        End If
        If G_P_This = 1 Then
            showContent
            Response.Write oblog.showpage(True, True, "个问题")
        Else
            If (G_P_This - 1) * G_P_PerMax < G_P_AllRecords Then
                rs.Move (G_P_This - 1) * G_P_PerMax
                Dim bookmark
                bookmark = rs.bookmark
                showContent
                Response.Write oblog.showpage(True, True, "个问题")
            Else
                G_P_This = 1
                showContent
                Response.Write oblog.showpage(True, True, "个问题")
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
		
End Sub 
Sub showcontent()
Dim i
i = 0
%>
	<div id="main_body">
		<ul class="main_top">
			<li class="main_top_left left"><%=sGuide%></li>
			<li class="main_top_right right"> </li>
		</ul>
		
			<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  
<style type="text/css">
<!--
.border tr td {padding:3px 0!important;}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="title">
    
    <td align="center" width="44"><strong>ID</strong></td>
    <td align="center" ><strong>问题</strong></td>
    <td align="center" width="300"><strong>答案</strong></td>
    

    <td align="center" width="100"><strong>操作</strong></td>
  </tr>
		
<%
	do while not rs.EOF
	Response.write "<tr align=""center""><td>"&rs(0)&"</td><td  align=""left"">&nbsp;&nbsp;&nbsp;&nbsp;"&rs(1)&"</td><td>"&rs(2)&"</td><td>  <A HREF=""?Action=edit&id="&rs(0)&""">修改</A>  |  <A HREF=""?Action=del&id="&rs(0)&""">删除</A>"
	i = i + 1
    If i >= G_P_PerMax Then Exit Do
    rs.movenext
Loop	
Response.write "</table>"
End Sub 
%>