<!--#include file="inc/inc_sys.asp"-->
<%
'----------------------------------------------
'合并原来的
'admin_regtext.asp/admin_userplacard.asp/admin_placard.asp/admin_friendsite.asp
'四个只操作oblog_setup表中单一字段的文件
'do-1:regtext;2:userolacard;3:placard;4:friendsite
'save-1/2/3/4
'URL中不允许直接传递字段名称等敏感数据
'----------------------------------------------
Dim Action,ActionId,ActionText,ActionField
Dim rs,strNote,strField
Action=LCase(Request.QueryString("Action"))
ActionId=Right(Action,1)
Action=Left(Action,Len(Action)-1)
Select Case ActionId
		Case "1"
			ActionText="修改友情连接(htm代码)"
			ActionField="site_friends"
		Case "2"
			ActionText="修改网站公告(htm代码)"
			ActionField="site_placard"
		Case "3"
			ActionText="修改注册条款(htm代码)"
			ActionField="reg_text"
		Case "4"
			ActionText="修改用户管理后台通知(htm代码)"
			ActionField="user_placard"
		Case Else
			Response.Write "错误的操作方式！"
			Response.End
End Select

if Action="saveconfig" then
	strNote=request("note")
	if not IsObject(conn) then link_database
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select "& ActionField &" from oblog_setup",conn,1,3
	rs(0)=strNote
	rs.update
	rs.close
	oblog.reloadsetup
	EventLog "进行修改友情链接、网站公告、后台通知、注册条款（文本方式）的操作!",oblog.NowUrl&"?"&Request.QueryString
	Set oBlog=Nothing
	response.Redirect("admin_note.asp?action=do" & ActionId)
else
	set rs=oblog.execute("select "& ActionField &" from oblog_setup")
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
		<li class="main_top_left left"><%=ActionText%></li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr>
    <td><form name="form1" method="post" action="admin_note.asp?Action=saveconfig<%=ActionId%>">
                <textarea name="note" cols="100" rows="25" id="edit"><%=rs(0)%></textarea>
				<br>
                <br>
                <input type="submit" name="Submit" value="提交修改">
      </form></td>
  </tr>
</table>
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
set rs=nothing
end if
%>