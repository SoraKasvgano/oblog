<!--#include file="inc/inc_sys.asp"-->
<%If CheckAccess("r_skin_user")=False Then Response.Write "无权操作":Response.End%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.form2.chkAll.checked){
	document.form2.chkAll.checked = document.form2.chkAll.checked&0;
    }
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
</SCRIPT>
<%
Dim ispass
Action=Trim(Request("Action"))

select case Action
	case "saveconfig"
		call saveconfig()
	case "showskin"
		call showskin()
	case "modiskin"
		call modiskin()
	case "savedefault"
		call savedefault()
	case "delconfig"
		call delconfig()
	case "addskin"
		call addskin()
	case "saveaddskin"
		call saveaddskin()
	case "passskin"
		call passskin
	case "unpassskin"
		call unpassskin
	Case Else
	Call showskin()
end select

sub showskin()
	dim rs,psql,sql,strField,userskinname,skinauthor,skinid,keyword,classid
ispass=CLng(Request("ispass"))
If ispass="" Or isnull(ispass) Then ispass=1
strField=trim(request("field"))
skinid=CLng(request("skinid"))
userskinname=Trim(request("userskinname"))
skinauthor=Trim(request("skinauthor"))
keyword=Trim(Request("keyword"))
classid=Request("classid")
if keyword<>"" then
	keyword=oblog.filt_badstr(keyword)
end if
G_P_FileName="m_userskin.asp?action=showskin"
if ispass=1 then
	G_P_FileName=G_P_FileName&"&ispass=1"
	psql=" and ispass=1 "
ElseIf ispass=0 then
	G_P_FileName=G_P_FileName&"&ispass=0"
	psql="and  ispass=0 "
end If
Select Case strField
	Case "userskinname"
		G_P_FileName=G_P_FileName&"&userskinname="&userskinname
		psql=" and userskinname  like '%" & Keyword & "%' "
	Case "skinauthor"
		G_P_FileName=G_P_FileName&"&skinauthor="&skinauthor
		psql=" and skinauthor like '%" & Keyword & "%' "
	Case "skinid"
		G_P_FileName=G_P_FileName&"&skinid="&skinid
		psql="and id="&clng(Keyword)
End Select
If classid<>"" Then
	If classid=0 Then
		G_P_FileName=G_P_FileName&"&classid="&classid
		psql=" And (classid=0 Or classid Is Null) "
	Else
		G_P_FileName=G_P_FileName&"&classid="&classid
		psql=" And classid=" & classid
	End If
End If

	if Request("page")<>"" then
	    G_P_This=cint(Request("page"))
	else
		G_P_This=1
	end if
	set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select id,userskinname,skinauthor,skinauthorurl,isdefault,ispass,skinpic from oblog_userskin  where 1=1 "&psql&" order by id desc "
		rs.Open sql,Conn,1,1
	  	if rs.eof and rs.bof then
            showContent(rs)
			G_P_Guide=G_P_Guide & " (共有0个模板)</h1>"
			Response.write "<div align='right'>"&G_P_Guide&"</div>"
		else
	    	G_P_AllRecords=rs.recordcount
			G_P_Guide=G_P_Guide & " (共有" & G_P_AllRecords & "个模板)</h1>"
			if G_P_This<1 then
	       		G_P_This=1
	    	end if
	    	if (G_P_This-1)*G_P_PerMax>G_P_AllRecords then
		   		if (G_P_AllRecords mod G_P_PerMax)=0 then
		     		G_P_This= G_P_AllRecords \ G_P_PerMax
			  	else
			      	G_P_This= G_P_AllRecords \ G_P_PerMax + 1
		   		end if
	    	end if
		    if G_P_This=1 then
	        	Call showContent(rs)
	        	Response.write oblog.showpage(true,true,"个模板")
	   	 	else
	   	     	if (G_P_This-1)*G_P_PerMax<G_P_AllRecords then
	         	   	rs.move  (G_P_This-1)*G_P_PerMax
	         		dim bookmark
	           		bookmark=rs.bookmark
	        	else
		        	G_P_This=1
		    	end if
		    	Call showContent(rs)
		    	Response.write oblog.showpage(true,true,"个模板")
			end if
		end if
end sub

sub showContent(rs)
	dim i
	i=0
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">用 户 模 板 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
<form name="form1" action="m_userskin.asp?action=showskin&ispass=1" method="post">
    <tr class="tdbg">
      <td width="100" height="30"><strong>按分类过滤：</strong></td>
      <td width="687" height="30">
      	<select size=1 name="classid">
      	  <option value="0">------尚未分类------</option>
          <%=sClasses%>
        </select>
        <input type="submit" value=" 查 看 "></td>
    </tr>
  </form>
<form name="form3" method="post" action="m_userskin.asp">
  <tr class="tdbg">
      <td width="120">
		<strong>高级查询：</strong> </td>
    <td ><select name="Field" id="Field">
	  <option value="userskinname" selected>按模板名</option>
	  <option value="skinauthor" selected>按作者名</option>
      <option value="skinid" >按模板ID</option>
      </select>
	   <input name="ispass" type="hidden" id="ispass" value="<%=ispass%>">
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 "></td>
  </tr>
</form>
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
		<li class="main_top_left left">用户模板管理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td  height=25 class="topbg" align="left"><strong>用户模板管理　　　<a href="m_userskin.asp?action=showskin&ispass=1">&gt;&gt;通过审核的模板</a>　　<a href="m_userskin.asp?action=showskin&ispass=0">&gt;&gt;未通过审核的模板</a></strong>
  </tr>
</table>
<form name="form2" method="post" action="m_userskin.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td height="25" colspan="6" ><strong><%if ispass=1 then Response.Write "通过审核的模板" else Response.write "未通过审核的模板"%></strong></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
  <tr class="topbg">
    <td align="center"><strong>选中</strong></td>
    <td align="center"><strong>模板名称</strong></td>
    <td><strong>操作</strong></td>
  </tr>
    <%
do while not rs.eof
dim userskinname
    userskinname=rs("userskinname")
%>
  <tr class="topbg">
    <td width="30" valign="top">
	<div align="center"><input name="checkbox" type="checkbox" onClick="unselectall()" id= "checkbox" class="tdbg" value='<%=rs("id")%>'></div>
	<div align="center"><%= rs("id") %></div></td>
    <td width="140"><div align="center">
	  <a href="../showskin.asp?id=<%=rs("id")%>" target="_blank"><img style="width:120px;height:77px;border:1px #888 solid;" src="<%=ProIco(rs("skinpic"),3)%>" /><br />
	  <%if rs("isdefault")=1 then
	  Response.Write "<span style=""color:#f00;font-weight:600;"">默认模板："&userskinname&"</span>"
	  else
	  Response.Write userskinname
	  end if
	  %>
	  </a>
	  </div></td>
    <td><div>
	  <strong>模板作者：</strong>
	  <%if rs("skinauthorurl")="" or isnull(rs("skinauthorurl")) then
	  Response.Write rs("skinauthor")
	  else
	  Response.Write "<a href="""&oblog.filt_html(rs("skinauthorurl"))&""" target='_blank'>"&rs("skinauthor")&"</a>"
	  end if%>
	  </div>
	  <div><%if rs("ispass")=1 then Response.Write("<span style=""color:#317531;font-weight:600;"">已审核</span>") else Response.Write("<span style=""color:#F30;font-weight:600;"">未审核</span>")%>
	<%if ispass=0 then%>
	<a href="m_userskin.asp?action=passskin&id=<%=rs("id")%>">通过审核</a>
	<%else%>
	<a href="m_userskin.asp?action=unpassskin&id=<%=rs("id")%>">取消审核</a>
	<%end if%>
	</div>
	  <div><a href="../admin_edit.asp?action=modiskin&skintype=user&t=0&editm=1&skinorder=0&id=<%=rs("id")%>"  target="_blank">修改主模板</a> <a href="../admin_edit.asp?action=modiskin&skintype=user&t=0&editm=1&skinorder=1&id=<%=rs("id")%>"  target="_blank">修改副模板</a> <a href="m_userskin.asp?action=modiskin&id=<%=rs("id")%>">修改模板(文本方式)</a></div>
	  <div><a href="m_userskin.asp?action=delconfig&id=<%=rs("id")%>" style="color:#f00;font-weight:600;" onclick=return(confirm("确定要删除这个模板吗？"))>删除模板</a></div>
	  </td>
    </tr>
    <%
	i=i+1
	if i>=G_P_PerMax then exit do
	rs.movenext
loop
%>
</table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="40" colspan="6" align="center" class="tdbg"> <div align="center">
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" />
	  全选
	 <input type="radio" value="savedefault" name="action" checked>默认模板</option>
	 <%if ispass=0 then%>
	  <input type="radio" value="passskin" name="action" >通过审核</option>
	  <%else%>
	  <input type="radio" value="unpassskin" name="action">取消审核</option>
	  <%end if%>
	   <input type="radio" value="delconfig" name="action" >删除</option>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="submit" name="Submit" value="保存设置">
        </div></td>
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
	set rs=nothing
end sub

Sub ReCountSkins()
	Dim rst,rst1
	Set rst=Server.CreateObject("Adodb.Recordset")
	Set rst1=Server.CreateObject("Adodb.Recordset")
	'重新计数
	rst.Open "select classid From oblog_skinclass",conn,1,3
	rst1.Open "select Count(id) ,Classid From oblog_userskin Where ispass=1 Group By classid",conn,1,3
	Do While Not rst.Eof
		rst1.Filter="classid=" & rst(0)
		If Not rst1.Eof Then
			oblog.Execute "Update oblog_skinclass Set icount=" & rst1(0) & " Where classid=" & rst(0)
		Else
			oblog.Execute "Update oblog_skinclass Set icount=0 Where classid=" & rst(0)
		End If
		rst.MoveNext
	Loop
	Set rst=Nothing
	Set rst1=Nothing
End Sub

sub savedefault()
	If Request.QueryString <>"" Then Exit Sub
	dim isdefaultID
	isdefaultID=Trim(Request("checkbox"))
	if instr(isdefaultID,",")>0 then
		oblog.showMsg "用户默认模板只可以选择一个！",""
	elseif isdefaultID="" then
		oblog.showMsg "请指定要设定为默认的模板！",""
	end if
	oblog.execute("update oblog_userskin set isdefault=0")
	oblog.execute("update oblog_userskin set isdefault=1 where id="&isdefaultID)
	WriteSysLog "进行了设定默认用户模板操作，目标模板ID："&isdefaultID&"",""
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""修改成功！"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub passskin()
	dim id
	id=Trim(Request("checkbox"))
	if instr(id,",")>0 then
	id=Replace(id," ","")
	oblog.execute("update oblog_userskin set ispass=1 where id in ("&id&")")
	elseif id="" then
	id=CLng(Request("id"))
		oblog.execute("update oblog_userskin set ispass=1 where id="&id)
	else
    	oblog.execute("update oblog_userskin set ispass=1 where id="&id)
	end if
	ReCountSkins
	WriteSysLog "进行了通过审核用户模板操作，目标模板ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "通过审核成功",""
end sub

sub unpassskin()
	dim id
	id=Trim(Request("checkbox"))
	if instr(id,",")>0 then
	id=Replace(id," ","")
	oblog.execute("update oblog_userskin set ispass=0 where id in ("&id&")")
	elseif id="" then
		id=CLng(Request("id"))
		oblog.execute("update oblog_userskin set ispass=0 where id="&id)
	else
		oblog.execute("update oblog_userskin set ispass=0 where id="&id)
	end if
	ReCountSkins
	WriteSysLog "进行了取消审核用户模板操作，目标模板ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "取消审核成功",""
end sub


sub saveconfig()
	dim rs,sql
	if Trim(Request("userskinname"))="" then oblog.sys_err("模板名不能为空"):Response.End()
	if Trim(Request("skinmain"))="" then oblog.sys_err("主模板不能为空"):Response.End()
	if Trim(Request("skinshowlog"))="" then oblog.sys_err("副模板不能为空"):Response.End()
	set rs=Server.CreateObject("adodb.recordset")
	sql="select * from oblog_userskin where id="&CLng(Request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs("userskinname")=Trim(Request("userskinname"))
	rs("skinauthor")=Trim(Request("skinauthor"))
	rs("skinmain")=Request("skinmain")
	rs("skinshowlog")=Request("skinshowlog")
	rs("skinpic")=Trim(Request("skinpic"))
	rs("skinauthorurl")=Trim(Request("skinauthorurl"))
	rs.update
	rs.close
	set rs=nothing
	ReCountSkins
	WriteSysLog "进行了修改用户模板操作（文本方式），目标模板ID："&Request.QueryString("id")&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "保存成功",""
end sub

sub delconfig()
    dim id
	id=Trim(Request("checkbox"))
	if instr(id,",")>0 then
	id=Replace(id," ","")
	oblog.execute("delete from oblog_userskin where id in ("&id&")")
	elseif id="" then
	id=CLng(Request.QueryString("id"))
	oblog.execute("delete from oblog_userskin where id="&id)
	else
	oblog.execute("delete from oblog_userskin where id="&id)
	end if
	ReCountSkins
	WriteSysLog "进行了删除用户模板操作，目标模板ID："&id&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "删除成功",""
end sub
sub modiconfig()
	dim rs
	set rs=oblog.execute("select * from oblog_userskin where id="&CLng(Request.QueryString("id")))
End Sub
sub saveaddskin()
	dim rs,sql
	set rs=Server.CreateObject("adodb.recordset")
	if Trim(Request("userskinname"))="" then oblog.sys_err("模板名不能为空"):Response.End()
	if Trim(Request("skinmain"))="" then oblog.sys_err("主模板不能为空"):Response.End()
	if Trim(Request("skinshowlog"))="" then oblog.sys_err("副模板不能为空"):Response.End()
	sql="select * from oblog_userskin where id="&CLng(Request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs.addnew
	rs("userskinname")=Trim(Request("userskinname"))
	rs("skinauthor")=Trim(Request("skinauthor"))
	rs("skinmain")=Trim(Request("skinmain"))
	rs("skinshowlog")=Trim(Request("skinshowlog"))
	rs("skinpic")=Trim(Request("skinpic"))
	rs("skinauthorurl")=Trim(Request("skinauthorurl"))
	rs.update
	rs.close
	set rs=nothing
	ReCountSkins
	WriteSysLog "进行了添加用户模板操作",oblog.NowUrl&"?"&Request.QueryString
	Response.Redirect "m_userskin.asp?action=showskin"
end sub

sub modiskin()
	dim rs
	set rs=oblog.execute("select * from oblog_userskin where id="&CLng(Request.QueryString("id")))
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">修改用户模板</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
    <tr class="tdbg">
    <td width="253" height="30"><strong>现在修改的模板是：<%=rs("userskinname")%></strong></td>
    <td width="516" height="30">
	<a href="m_userskin.asp?action=modiskin&id=<%=rs("id")%>">修改模板</a>　　<a href="m_userskin.asp?action=showskin&ispass=1">返回管理菜单</a>
      <a href="m_skin_help.asp" target="_blank"><strong>模板标记帮助</strong></a></td>
    </tr>
</table>

<form method="POST" action="m_userskin.asp?id=<%=CLng(Request.QueryString("id"))%>" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="769" height="22" class="topbg"><strong>修改模板</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">模板名称：
        <input name="userskinname" type="text" id="userskinname" value=<%=rs("userskinname")%>>
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>>
        <br>
        作者连接：
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="<%=rs("skinauthorurl")%>">
         <br>
        预览图片<strong>：
        <input name="skinpic" type="text" id="skinpic" size="50" value="<%=rs("skinpic")%>">
        </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <strong>主模板：</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"><%if rs("skinmain")<>"" then Response.Write Server.HtmlEncode(rs("skinmain")) else Response.Write("")%></textarea>
        <br>
        <br>
        <strong>副模板： <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"><%if rs("skinshowlog")<>"" then Response.Write Server.HtmlEncode(rs("skinshowlog")) else Response.Write("")%></textarea>
        </strong></td>
    </tr>
    <tr>
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存修改 " >
      </div></td>
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
set rs=nothing
end sub

sub addskin()
%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">添加用户模板</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
  <tr class="tdbg">
    <td height="30"><div align="center"><a href="m_userskin.asp?action=showskin"><strong>返回管理菜单</strong></a>　　 <a href="m_skin_help.asp" target="_blank"><strong>模板标记帮助</strong></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</div></td>
  </tr>
</table>

<form method="POST" action="m_userskin.asp" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>模板参数</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">模板名称：
        <input name="userskinname" type="text" id="userskinname">
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor">
        <br>
        作者连接<strong>：
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="">
        </strong> <br>
        预览图片<strong>：
        <input name="skinpic" type="text" id="skinpic" size="50">
        </strong> </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <strong>主模板：</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"></textarea>
        <br>
        <br>
        <strong>副模板： <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"></textarea>
        </strong></td>
    </tr>
    <tr>
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveaddskin">
          <input name="cmdadd" type="submit" id="cmdadd" value=" 添加 " >
      </div></td>
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
end Sub
Function sClasses()
Dim rstClass
Set rstClass=Server.CreateObject("Adodb.RecordSet")
rstClass.Open "select * From oblog_skinclass Where iType=0",conn,1,3
If Not rstClass.Eof Then
	Do While Not rstClass.Eof
		sClasses= sClasses & "<option value=" & rstClass("classid") & " >" & rstClass("classname") & "(" & rstClass("icount") & ")</option>" & vbcrlf
		rstClass.MoveNext
	Loop
	rstClass.MoveFirst
End if
End Function 
Set oblog = Nothing
%>