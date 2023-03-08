<!--#include file="inc/inc_sys.asp"-->
<%
If CheckAccess("r_words")=False Then Response.Write "无权操作":Response.End
Action=Trim(Request("Action"))
if Action="saveconfig" then
	call saveconfig()
else
	call showconfig()
end if

sub showconfig()
	Dim rs,badstr1,badstr2,badstr3,badstr4
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "select * from oblog_config Where id in(6,7,8,9)",conn,1,3

	rs.Filter="id=6"
	If Not rs.Eof Then
		badstr1=OB_IIF(rs("ob_value"),"")
	End If
	rs.Filter="id=7"
	If Not rs.Eof Then
		badstr2=OB_IIF(rs("ob_value"),"")
	End If
	rs.Filter="id=8"
	If Not rs.Eof Then
		badstr3=OB_IIF(rs("ob_value"),"")
	End If
	rs.Filter="id=9"
	If Not rs.Eof Then
		badstr4=OB_IIF(rs("ob_value"),"")
	End If
	rs.Close
	Set rs=Nothing

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--后台管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">关 键 字 管 理</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="m_words.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"> <strong>禁止发表的关键字(封杀关键字)：</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">当评论，日志，留言，标签(tag)中含有以下关键字将被禁止发表，请您将要禁止发表的字符串添入，如果有多个字符串，请用回车分隔开。<br>
        可以输入恶意网址来禁止垃圾评论。</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <textarea name="badstr1" cols="35" rows="10" id="badstr1">
<%=badstr1%></textarea>      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg">&nbsp; </td>
    </tr>
    <tr>
    <tr>
      <td height="22" class="topbg"> <strong>日志敏感字过滤(可疑关键字)：</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">此处设置影响到日志，评论，专题名字、blog名字、模板设置的过滤。内容中出现关键字后，如果是日志则被设置为可疑，如果是其他内容则禁止保存。</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <textarea name="badstr2" cols="35" rows="10" id="badstr2">
<%=badstr2%></textarea>      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg">&nbsp;</td>
    </tr>
	 <tr>
      <td height="22" class="topbg"> <strong>关键字替换：</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">此处设置影响到日志，评论，专题名字、blog名字、模板设置的过滤。过滤字符将过滤内容中包含以下字符的内容(以×号代替)，请您将要过滤的字符串添入，如果有多个字符串，请用回车分隔开。</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <textarea name="badstr3" cols="35" rows="10" id="badstr3">
<%=badstr3%></textarea>      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg">&nbsp;</td>
    </tr>
      <td height="25" class="topbg"><strong>注册过滤字符</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <p>注册过滤字符将不允许用户注册包含以下字符的内容，请您将要过滤的字符串添入，如果有多个字符串，请用回车隔开。<br>
      	注册过滤字符已经包含前面所设置的3项中的关键字,此处不必重复填写
        </p></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"> <p class="tdbg">
          <textarea name="badstr4" cols="35" rows="10" id="reg_badstr"><%=badstr4%></textarea>
        </p></td>
    </tr>
    <tr>
      <td height="40" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " > </td>
    </tr>
  </table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</form>
</body>
</html>
<%
set rs=nothing
end sub

sub saveconfig()
	If Request.QueryString <> "" Then Exit Sub
	Dim rs,badstr1,badstr2,badstr3,badstr4
	if not IsObject(conn) then link_database
	badstr1=oblog.FilterEmpty(Request("badstr1"))
	badstr2=oblog.FilterEmpty(Request("badstr2"))
	badstr3=oblog.FilterEmpty(Request("badstr3"))
	badstr4=oblog.FilterEmpty(Request("badstr4"))
	set rs=Server.CreateObject("adodb.recordset")

	rs.open "select * from oblog_config Where id=6",conn,1,3
	If  rs.Eof Then rs.AddNew:rs("id")=6
	rs("ob_value")=badstr1
	rs.Update
	rs.Close

	rs.open "select * from oblog_config Where id=7",conn,1,3
	If  rs.Eof Then rs.AddNew:rs("id")=7
	rs("ob_value")=badstr2
	rs.Update
	rs.Close

	rs.open "select * from oblog_config Where id=8",conn,1,3
	If  rs.Eof Then rs.AddNew:rs("id")=8
	rs("ob_value")=badstr3
	rs.Update
	rs.Close

	rs.open "select * from oblog_config Where id=9",conn,1,3
	If  rs.Eof Then rs.AddNew:rs("id")=9
	rs("ob_value")=badstr4
	rs.Update
	rs.Close
	set rs=nothing
	oblog.ReloadCache
	WriteSysLog "进行了关键字管理操作",""
'	Response.Redirect "m_words.asp"
	Oblog.ShowMsg "操作成功",""
end Sub
Set oblog = Nothing
%>