<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/Cls_XmlDoc.asp"-->
<%
dim action,rs,site_placard
Dim ADPath
Action=Trim(Request("Action"))
Dim xmlDoc
Set xmlDoc = New Cls_XmlDoc
xmlDoc.Unicode = False
ADPath = blogdir & oblog.CacheConfig(80)
If Not xmlDoc.LoadXml (ADPath& "/GG.config") Then
	If xmlDoc.LoadXml (blogdir& "XmlData/GG.config") Then
		xmlDoc.SaveAs (ADPath& "/GG.config")
	Else
		Response.Write (blogdir&"XmlData/GG.config 不存在，无法继续操作！")
		Set XmlDoc = Nothing
		Response.End
	End If
End If
xmlDoc.LoadXml (ADPath& "/GG.config")
if action="saveconfig" Then
	If Request.QueryString <>"" Then Response.End
	dim fso,dirstr,gg_usertop,gg_usercomment,gg_userbot,gg_userlinks,gg_user_desktop_main
	gg_usertop=Trim(Request("gg_usertop"))
	gg_usercomment=Trim(Request("gg_usercomment"))
	gg_userbot=Trim(Request("gg_userbot"))
	gg_userlinks=Trim(Request("gg_userlinks"))
	gg_user_desktop_main=Trim(Request("gg_user_desktop_main"))
	set fso=Server.CreateObject(oblog.CacheCompont(1))
	dirstr=Server.MapPath(ADPath)
	if fso.FolderExists(dirstr)=false then fso.CreateFolder(dirstr)
	Call oblog.BuildFile(dirstr&"\gg_usercomment.htm",gg_usercomment)
	Call oblog.BuildFile(dirstr&"\gg_usertop.htm",gg_usertop)
	Call oblog.BuildFile(dirstr&"\gg_userbot.htm",gg_userbot)
	Call oblog.BuildFile(dirstr&"\gg_userlinks.htm",gg_userlinks)
	Call oblog.BuildFile(dirstr&"\gg_user_desktop_main.htm",gg_user_desktop_main)
	xmlDoc.UpdateNodeText "gg_usertop",oblog.htm2js_Script(gg_usertop,"gg_usertop"),True
	xmlDoc.UpdateNodeText "gg_userbot",oblog.htm2js_Script(gg_userbot,"gg_userbot"),True
	xmlDoc.UpdateNodeText "gg_userlinks",oblog.htm2js_Script(gg_userlinks,"gg_userlinks"),True
	xmlDoc.UpdateNodeText "gg_usercomment",oblog.htm2js_Script(gg_usercomment,"gg_usercomment"),True
	xmlDoc.UpdateNodeText "gg_user_desktop_main",oblog.htm2js_Script(gg_user_desktop_main,"gg_user_desktop_main"),True
	'兼容旧广告
	Call oblog.BuildFile(dirstr&"\ad_usercomment.htm",gg_usercomment)
	Call oblog.BuildFile(dirstr&"\ad_usertop.htm",gg_usertop)
	Call oblog.BuildFile(dirstr&"\ad_userbot.htm",gg_userbot)
	Call oblog.BuildFile(dirstr&"\ad_userlinks.htm",gg_userlinks)
	xmlDoc.UpdateNodeText "ad_usertop",oblog.htm2js_Script(gg_usertop,"ad_usertop"),True
	xmlDoc.UpdateNodeText "ad_userbot",oblog.htm2js_Script(gg_userbot,"ad_userbot"),True
	xmlDoc.UpdateNodeText "ad_userlinks",oblog.htm2js_Script(gg_userlinks,"ad_userlinks"),True
	xmlDoc.UpdateNodeText "ad_usercomment",oblog.htm2js_Script(gg_usercomment,"ad_usercomment"),True
	xmlDoc.Save
	set fso=Nothing
	Set XmlDoc = Nothing
	EventLog "进行修改用户页面广告的操作!",""
	oblog.ShowMsg "保存完成",""
else

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>blog页面广告设置</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">修改用户blog页面广告</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">

<form name="form1" method="post" action="admin_ad.asp">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
          <tr>
            <td>
            </td>
          </tr>
    <tr>
      <td height="22" class="topbg"><strong>用户页面顶部广告（此段代码显示在所有blog用户页面顶部，可以放置弹出广告等）</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">
                <textarea name="gg_usertop" cols="80" rows="8" ><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_usertop.htm")%></textarea>
          </td></tr>
  </table>
  <br />
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>用户页面底部广告（此段代码显示在所有blog用户页面最底部，可以书写版权信息等代码）</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">
                <textarea name="gg_userbot" cols="80" rows="8" ><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_userbot.htm")%></textarea>
          </td></tr>
  </table>
  <br />
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>用户回复上部广告（此段代码显示在所有blog用户评论窗口上部）</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">
                <textarea name="gg_usercomment" cols="80" rows="8"><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_usercomment.htm")%></textarea>
          </td></tr>
  </table>

<br />
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>用户友情连接部分广告（此段代码显示在所有blog用户友情连接标签部分）</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">

                <textarea name="gg_userlinks" cols="80" rows="8"><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_userlinks.htm")%></textarea>
</td>
    </tr>
  </table>
    <br />
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>用户后台公告下部广告</strong><br/>（此段代码显示在所有blog用户后台公告下部,大小330*100.请不要加过多js代码以免不好的用户体验.）</td>
    </tr>
    <tr>
      <td height="25" class="tdbg">
                <textarea name="gg_user_desktop_main" cols="80" rows="8"><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_user_desktop_main.htm")%></textarea>
          </td></tr>
    <tr>
      <td height="40" align="center"><input name="Action" type="hidden" id="Action" value="saveconfig">
                <input type="submit" name="Submit" value="提交修改">
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
</body>
</html>
<%end If
Set oblog=Nothing
%>